package com.example.editdocx;

import com.aspose.words.Comment;
import com.aspose.words.CommentRangeEnd;
import com.aspose.words.CommentRangeStart;
import com.aspose.words.Document;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.SaveFormat;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.UUID;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class WordCommentService {
    private static final Pattern BODY_PATTERN = Pattern.compile("<body[^>]*>(.*?)</body>", Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
    private final Path uploadDir = Path.of("uploads");

    public WordCommentService() throws Exception {
        Files.createDirectories(uploadDir);
    }

    public UploadPreview saveAndPreview(MultipartFile file) throws Exception {
        validateDocx(file);

        String uploadId = UUID.randomUUID().toString().replace("-", "").substring(0, 12);
        Path workDir = uploadDir.resolve(uploadId);
        Files.createDirectories(workDir);

        Path originalPath = workDir.resolve("original.docx");
        file.transferTo(originalPath);

        String html = convertDocxToHtml(originalPath);
        return new UploadPreview(uploadId, originalFilename(file), html);
    }

    public GeneratedDocx generateFromUpload(String uploadId, List<CommentDto> comments) throws Exception {
        Path workDir = uploadDir.resolve(uploadId);
        Path originalPath = workDir.resolve("original.docx");
        if (!Files.exists(originalPath)) {
            throw new IllegalArgumentException("上传文件不存在，请重新上传 DOCX");
        }

        return generate(originalPath, workDir.resolve("java-aspose-comments.docx"), "java-aspose-comments.docx", comments);
    }

    public GeneratedDocx generateFromFile(MultipartFile file, List<CommentDto> comments) throws Exception {
        validateDocx(file);

        Path workDir = uploadDir.resolve(UUID.randomUUID().toString().replace("-", "").substring(0, 12));
        Files.createDirectories(workDir);

        Path originalPath = workDir.resolve("original.docx");
        file.transferTo(originalPath);

        return generate(originalPath, workDir.resolve("java-aspose-comments.docx"), "java-aspose-comments.docx", comments);
    }

    private GeneratedDocx generate(Path inputPath, Path outputPath, String downloadName, List<CommentDto> comments) throws Exception {
        if (comments == null || comments.isEmpty()) {
            throw new IllegalArgumentException("至少需要一条批注");
        }

        Document document = new Document(inputPath.toString());
        ApplyResult result = applyComments(document, comments);
        document.save(outputPath.toString());
        return new GeneratedDocx(outputPath, downloadName, result.applied(), result.skipped());
    }

    private String convertDocxToHtml(Path inputPath) throws Exception {
        Document document = new Document(inputPath.toString());
        Path htmlPath = inputPath.resolveSibling("preview.html");

        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportImagesAsBase64(true);
        options.setPrettyFormat(true);
        document.save(htmlPath.toString(), options);

        String html = Files.readString(htmlPath, StandardCharsets.UTF_8);
        return extractBodyHtml(html);
    }

    private ApplyResult applyComments(Document document, List<CommentDto> comments) throws Exception {
        List<CommentDto> pending = comments.stream().map(this::normalizeComment).toList();
        NodeCollection paragraphNodes = document.getChildNodes(NodeType.PARAGRAPH, true);
        List<Paragraph> paragraphs = new ArrayList<>();
        for (Object item : paragraphNodes) {
            Node node = (Node) item;
            paragraphs.add((Paragraph) node);
        }

        List<ParagraphComments> paragraphComments = new ArrayList<>();
        for (CommentDto comment : pending) {
            LocatedComment located = locateComment(paragraphs, comment);
            if (located == null) {
                continue;
            }
            ParagraphComments bucket = paragraphComments.stream()
                .filter(item -> item.paragraph() == located.paragraph())
                .findFirst()
                .orElseGet(() -> {
                    ParagraphComments created = new ParagraphComments(located.paragraph(), new ArrayList<>());
                    paragraphComments.add(created);
                    return created;
                });
            bucket.comments().add(located);
        }

        int applied = 0;
        for (ParagraphComments item : paragraphComments) {
            applied += rebuildParagraphWithComments(document, item.paragraph(), item.comments());
        }

        return new ApplyResult(applied, pending.size() - applied);
    }

    private CommentDto normalizeComment(CommentDto comment) {
        CommentDto normalized = new CommentDto();
        normalized.setSelectedText(normalizeText(nullToEmpty(comment.getSelectedText())));
        normalized.setOccurrence(comment.getOccurrence());
        normalized.setComment(nullToEmpty(comment.getComment()).trim());
        normalized.setAuthor(nullToEmpty(comment.getAuthor()).isBlank() ? "Reviewer" : comment.getAuthor().trim());
        normalized.setInitials(nullToEmpty(comment.getInitials()).isBlank() ? "RV" : comment.getInitials().trim());
        return normalized;
    }

    private LocatedComment locateComment(List<Paragraph> paragraphs, CommentDto comment) {
        if (comment.getSelectedText().isBlank() || comment.getComment().isBlank()) {
            return null;
        }

        int seen = 0;
        for (Paragraph paragraph : paragraphs) {
            String rawText = getParagraphRunText(paragraph);
            NormalizedText normalized = normalizeTextWithMapping(rawText);
            int searchFrom = 0;

            while (true) {
                int start = normalized.text().indexOf(comment.getSelectedText(), searchFrom);
                if (start < 0) {
                    break;
                }

                if (seen == comment.getOccurrence()) {
                    int originalStart = normalized.mapping().get(start);
                    int originalEnd = normalized.mapping().get(start + comment.getSelectedText().length() - 1) + 1;
                    return new LocatedComment(
                        paragraph,
                        originalStart,
                        originalEnd,
                        comment.getSelectedText(),
                        comment.getComment(),
                        comment.getAuthor(),
                        comment.getInitials()
                    );
                }

                seen++;
                searchFrom = start + comment.getSelectedText().length();
            }
        }
        return null;
    }

    private int rebuildParagraphWithComments(Document document, Paragraph paragraph, List<LocatedComment> locatedComments) throws Exception {
        String text = getParagraphRunText(paragraph);
        List<RunSpan> sourceRuns = collectRunSpans(paragraph);
        int cursor = 0;

        List<LocatedComment> filtered = new ArrayList<>();
        for (LocatedComment item : locatedComments.stream().sorted(Comparator.comparingInt(LocatedComment::start)).toList()) {
            if (item.start() < cursor) {
                continue;
            }
            filtered.add(item);
            cursor = item.end();
        }

        paragraph.removeAllChildren();
        int position = 0;
        for (LocatedComment item : filtered) {
            if (item.start() > position) {
                appendPreservedRuns(paragraph, sourceRuns, position, item.start());
            }

            appendCommentRange(document, paragraph, sourceRuns, text, item);
            position = item.end();
        }

        if (position < text.length()) {
            appendPreservedRuns(paragraph, sourceRuns, position, text.length());
        }
        return filtered.size();
    }

    private void appendCommentRange(Document document, Paragraph paragraph, List<RunSpan> sourceRuns, String text, LocatedComment item) throws Exception {
        Comment comment = new Comment(document, item.author(), item.initials(), new Date());
        comment.setText(item.comment());

        paragraph.appendChild(new CommentRangeStart(document, comment.getId()));
        List<Run> commentRuns = appendPreservedRuns(paragraph, sourceRuns, item.start(), item.end());
        if (commentRuns.isEmpty()) {
            paragraph.appendChild(new Run(document, text.substring(item.start(), item.end())));
        }
        paragraph.appendChild(new CommentRangeEnd(document, comment.getId()));
        paragraph.appendChild(comment);
    }

    private List<RunSpan> collectRunSpans(Paragraph paragraph) {
        List<RunSpan> spans = new ArrayList<>();
        int position = 0;
        for (Object item : paragraph.getChildNodes(NodeType.RUN, false)) {
            Node node = (Node) item;
            Run run = (Run) node;
            String text = run.getText();
            if (text == null || text.isEmpty()) {
                continue;
            }

            int start = position;
            int end = start + text.length();
            spans.add(new RunSpan(start, end, text, run));
            position = end;
        }
        return spans;
    }

    private List<Run> appendPreservedRuns(Paragraph paragraph, List<RunSpan> sourceRuns, int start, int end) throws Exception {
        List<Run> appendedRuns = new ArrayList<>();
        if (start >= end) {
            return appendedRuns;
        }

        for (RunSpan span : sourceRuns) {
            int overlapStart = Math.max(start, span.start());
            int overlapEnd = Math.min(end, span.end());
            if (overlapStart >= overlapEnd) {
                continue;
            }

            int relativeStart = overlapStart - span.start();
            int relativeEnd = overlapEnd - span.start();
            Run clonedRun = (Run) span.run().deepClone(true);
            clonedRun.setText(span.text().substring(relativeStart, relativeEnd));
            paragraph.appendChild(clonedRun);
            appendedRuns.add(clonedRun);
        }
        return appendedRuns;
    }

    private String getParagraphRunText(Paragraph paragraph) {
        StringBuilder builder = new StringBuilder();
        for (Object item : paragraph.getChildNodes(NodeType.RUN, false)) {
            Node node = (Node) item;
            builder.append(((Run) node).getText());
        }
        return builder.toString();
    }

    private String extractBodyHtml(String html) {
        Matcher matcher = BODY_PATTERN.matcher(html);
        return matcher.find() ? matcher.group(1).trim() : html;
    }

    private NormalizedText normalizeTextWithMapping(String value) {
        StringBuilder chars = new StringBuilder();
        List<Integer> mapping = new ArrayList<>();
        Integer pendingSpaceIndex = null;
        boolean emittedText = false;

        for (int index = 0; index < value.length(); index++) {
            char current = value.charAt(index);
            if (Character.isWhitespace(current)) {
                if (emittedText) {
                    pendingSpaceIndex = index;
                }
                continue;
            }

            if (pendingSpaceIndex != null) {
                chars.append(' ');
                mapping.add(pendingSpaceIndex);
                pendingSpaceIndex = null;
            }

            chars.append(current);
            mapping.add(index);
            emittedText = true;
        }

        return new NormalizedText(chars.toString(), mapping);
    }

    private String normalizeText(String value) {
        return value.trim().replaceAll("\\s+", " ");
    }

    private void validateDocx(MultipartFile file) {
        String filename = originalFilename(file);
        if (!filename.toLowerCase().endsWith(".docx")) {
            throw new IllegalArgumentException("仅支持 .docx 文件");
        }
    }

    private String originalFilename(MultipartFile file) {
        return file.getOriginalFilename() == null || file.getOriginalFilename().isBlank()
            ? "document.docx"
            : file.getOriginalFilename();
    }

    private String nullToEmpty(String value) {
        return value == null ? "" : value;
    }

    public record UploadPreview(String uploadId, String fileName, String html) {
    }

    public record GeneratedDocx(Path path, String downloadName, int applied, int skipped) {
    }

    private record ApplyResult(int applied, int skipped) {
    }

    private record ParagraphComments(Paragraph paragraph, List<LocatedComment> comments) {
    }

    private record LocatedComment(Paragraph paragraph, int start, int end, String text, String comment, String author, String initials) {
    }

    private record RunSpan(int start, int end, String text, Run run) {
    }

    private record NormalizedText(String text, List<Integer> mapping) {
    }
}
