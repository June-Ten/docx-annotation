from __future__ import annotations

import datetime as dt
import json
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path

import aspose.words as aw
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


@dataclass
class LocatedComment:
    """前端批注在某个 Aspose Paragraph 中的定位结果。"""

    start: int
    end: int
    text: str
    comment: str
    author: str
    initials: str


@app.post("/api/comments/generate")
def generate_comments_docx():
    """接收原始 DOCX 和批注列表，使用 Aspose.Words 生成带批注的新 DOCX。"""

    uploaded_file = request.files.get("file")
    if uploaded_file is None:
        return jsonify({"message": "请上传 DOCX 文件"}), 400

    # 前端用 multipart/form-data 传递 comments 字段，内容是 JSON 字符串。
    comments_payload = request.form.get("comments", "[]")
    try:
        comments = json.loads(comments_payload)
    except json.JSONDecodeError:
        return jsonify({"message": "批注数据不是合法 JSON"}), 400

    if not isinstance(comments, list) or len(comments) == 0:
        return jsonify({"message": "至少需要一条批注"}), 400

    # 先用原始文件名判断扩展名，避免中文文件名被 secure_filename 清空后误判。
    original_filename = uploaded_file.filename or "document.docx"
    if not original_filename.lower().endswith(".docx"):
        return jsonify({"message": "仅支持 .docx 文件"}), 400

    # 保存到临时目录时使用安全文件名，防止路径穿越或非法字符。
    safe_name = secure_filename(original_filename)
    safe_stem = Path(safe_name).stem or "document"
    filename = f"{safe_stem}.docx"

    # Windows 下下载流读取期间文件可能被锁定，所以等响应关闭后再清理目录。
    work_dir = Path(tempfile.mkdtemp(dir=UPLOAD_DIR))
    input_path = work_dir / filename
    output_path = work_dir / append_suffix(filename, "-aspose-comments")
    uploaded_file.save(input_path)

    document = aw.Document(str(input_path))
    applied, skipped = apply_comments(document, comments)
    document.save(str(output_path))

    response = send_file(
        output_path,
        as_attachment=True,
        download_name=output_path.name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    response.call_on_close(lambda: shutil.rmtree(work_dir, ignore_errors=True))
    response.headers["X-Comments-Applied"] = str(applied)
    response.headers["X-Comments-Skipped"] = str(skipped)
    return response


def apply_comments(document, comments: list[dict]) -> tuple[int, int]:
    """把前端批注定位到段落，并用 Aspose 的 CommentRangeStart/End 写入批注。"""

    pending = [normalize_comment(comment) for comment in comments]
    paragraph_map = {}

    # Aspose 可以深度获取正文、表格等位置中的所有段落。
    paragraphs = [
        node.as_paragraph()
        for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    ]

    for comment in pending:
        location = locate_comment(paragraphs, comment)
        if location is None:
            continue

        paragraph, located = location
        paragraph_map.setdefault(paragraph, []).append(located)

    applied = 0
    for paragraph, located_comments in paragraph_map.items():
        applied += rebuild_paragraph_with_comments(document, paragraph, located_comments)

    return applied, len(pending) - applied


def normalize_comment(comment: dict) -> dict:
    """规范化前端提交的批注字段，补齐默认作者信息。"""

    return {
        "selectedText": normalize_text(str(comment.get("selectedText", ""))),
        "occurrence": int(comment.get("occurrence", 0) or 0),
        "comment": str(comment.get("comment", "")).strip(),
        "author": str(comment.get("author", "Reviewer")).strip() or "Reviewer",
        "initials": str(comment.get("initials", "RV")).strip() or "RV",
    }


def locate_comment(paragraphs: list, comment: dict) -> tuple[object, LocatedComment] | None:
    """按“选中文本 + 第几次出现”在 Aspose 段落中定位批注。"""

    selected_text = comment["selectedText"]
    if not selected_text or not comment["comment"]:
        return None

    seen = 0
    for paragraph in paragraphs:
        raw_text = get_paragraph_run_text(paragraph)
        paragraph_text, mapping = normalize_text_with_mapping(raw_text)
        search_from = 0

        while True:
            start = paragraph_text.find(selected_text, search_from)
            if start == -1:
                break

            if seen == comment["occurrence"]:
                original_start = mapping[start]
                original_end = mapping[start + len(selected_text) - 1] + 1
                return (
                    paragraph,
                    LocatedComment(
                        start=original_start,
                        end=original_end,
                        text=selected_text,
                        comment=comment["comment"],
                        author=comment["author"],
                        initials=comment["initials"],
                    ),
                )

            seen += 1
            search_from = start + len(selected_text)

    return None


def rebuild_paragraph_with_comments(document, paragraph, located_comments: list[LocatedComment]) -> int:
    """重建段落内的 run，并在选中文本前后插入 Aspose 批注范围节点。"""

    text = get_paragraph_run_text(paragraph)
    source_runs = collect_run_spans(paragraph)
    ordered = sorted(located_comments, key=lambda item: item.start)
    non_overlapping: list[LocatedComment] = []
    cursor = 0

    # Word 批注锚点不能互相交叉；重叠批注先跳过，避免生成损坏文档。
    for item in ordered:
        if item.start < cursor:
            continue
        non_overlapping.append(item)
        cursor = item.end

    # 清空段落子节点后重建，段落级格式仍保留在 paragraph_format 中。
    paragraph.remove_all_children()
    position = 0
    applied = 0

    for item in non_overlapping:
        if item.start > position:
            append_preserved_runs(paragraph, source_runs, position, item.start)

        append_comment_range(document, paragraph, source_runs, text, item)
        position = item.end
        applied += 1

    if position < len(text):
        append_preserved_runs(paragraph, source_runs, position, len(text))

    return applied


def append_comment_range(document, paragraph, source_runs: list[dict], text: str, item: LocatedComment) -> None:
    """插入 CommentRangeStart/End，并把 Comment 节点挂在同一段落中。"""

    comment = aw.Comment(document, item.author, item.initials, dt.datetime.now())
    comment.set_text(item.comment)

    paragraph.append_child(aw.CommentRangeStart(document, comment.id))
    comment_runs = append_preserved_runs(paragraph, source_runs, item.start, item.end)
    if not comment_runs:
        paragraph.append_child(aw.Run(document, text[item.start:item.end]))
    paragraph.append_child(aw.CommentRangeEnd(document, comment.id))
    paragraph.append_child(comment)


def collect_run_spans(paragraph) -> list[dict]:
    """记录原段落每个 run 在整段文本中的起止位置和格式来源。"""

    spans = []
    position = 0
    for node in paragraph.get_child_nodes(aw.NodeType.RUN, False):
        run = node.as_run()
        text = run.text
        if not text:
            continue

        start = position
        end = start + len(text)
        spans.append({"start": start, "end": end, "text": text, "run": run})
        position = end

    return spans


def append_preserved_runs(paragraph, source_runs: list[dict], start: int, end: int) -> list:
    """按原 run 边界写入文本片段，并通过 clone 保留原 run 字体、字号等格式。"""

    appended_runs = []
    if start >= end:
        return appended_runs

    for span in source_runs:
        overlap_start = max(start, span["start"])
        overlap_end = min(end, span["end"])
        if overlap_start >= overlap_end:
            continue

        relative_start = overlap_start - span["start"]
        relative_end = overlap_end - span["start"]
        cloned_run = span["run"].clone(True).as_run()
        cloned_run.text = span["text"][relative_start:relative_end]
        paragraph.append_child(cloned_run)
        appended_runs.append(cloned_run)

    return appended_runs


def get_paragraph_run_text(paragraph) -> str:
    """只拼接 run 文本，避免 paragraph.get_text() 带入段落结束符等控制字符。"""

    return "".join(
        node.as_run().text
        for node in paragraph.get_child_nodes(aw.NodeType.RUN, False)
    )


def normalize_text(value: str) -> str:
    """折叠连续空白，降低 HTML 预览文本和 DOCX 原文之间的空白差异。"""

    return " ".join(value.split())


def normalize_text_with_mapping(value: str) -> tuple[str, list[int]]:
    """折叠空白并保留每个规范化字符对应的原始文本下标。"""

    chars = []
    mapping = []
    pending_space_index = None
    emitted_text = False

    for index, char in enumerate(value):
        if char.isspace():
            if emitted_text:
                pending_space_index = index
            continue

        if pending_space_index is not None:
            chars.append(" ")
            mapping.append(pending_space_index)
            pending_space_index = None

        chars.append(char)
        mapping.append(index)
        emitted_text = True

    return "".join(chars), mapping


def append_suffix(filename: str, suffix: str) -> str:
    """给文件名追加后缀，保留原扩展名。"""

    path = Path(filename)
    return f"{path.stem}{suffix}{path.suffix}"


@app.get("/api/health")
def health():
    return jsonify({"ok": True, "backend": "aspose-words"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
