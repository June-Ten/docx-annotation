# DOCX 划词批注

前端：`frontend`，Vue 3 + Vite + mammoth.js + web-highlighter。

后端：

- `backend-python`：Flask + python-docx，默认端口 `5000`。
- `backend-python-aspose`：Flask + aspose-words，默认端口 `5001`。

## 运行

```bash
cd backend-python
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Aspose 后端：

```bash
cd backend-python-aspose
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

```bash
cd frontend
npm install
npm run dev
```

默认地址：

- 前端：http://localhost:5173
- python-docx 后端：http://localhost:5000
- Aspose 后端：http://localhost:5001

前端已配置两条 Vite 代理：

- `/api/python-docx/*` -> `http://localhost:5000/api/*`
- `/api/aspose/*` -> `http://localhost:5001/api/*`

## 功能

- 前端上传 DOCX 后用 mammoth.js 转成 HTML 预览。
- 在预览区域划词后，web-highlighter 自动高亮选区。
- 保存批注后，可分别点击“生成批注文件（python-docx）”或“生成批注文件（Aspose）”，把原始 DOCX 和批注列表发送到不同后端。
- `backend-python` 用 python-docx 的 `Document.add_comment()` 在匹配文本上生成 Word 批注。
- `backend-python-aspose` 用 Aspose.Words 的 `CommentRangeStart`、`CommentRangeEnd`、`Comment` 生成 Word 批注。

## 说明

当前版本按 `选中文本 + 第几次出现` 定位原文。建议批注单段落内的文本；跨段落选择、复杂表格/页眉页脚内容可能需要进一步增强定位逻辑。
