package com.example.editdocx;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;

@CrossOrigin
@RestController
@RequestMapping("/api")
public class DocxController {
    private final ObjectMapper objectMapper;
    private final WordCommentService wordCommentService;

    public DocxController(ObjectMapper objectMapper, WordCommentService wordCommentService) {
        this.objectMapper = objectMapper;
        this.wordCommentService = wordCommentService;
    }

    @PostMapping(value = "/docx/upload", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public Map<String, String> uploadDocx(@RequestParam("file") MultipartFile file) throws Exception {
        WordCommentService.UploadPreview preview = wordCommentService.saveAndPreview(file);
        return Map.of(
            "uploadId", preview.uploadId(),
            "fileName", preview.fileName(),
            "html", preview.html()
        );
    }

    @PostMapping("/docx/{uploadId}/comments")
    public ResponseEntity<ByteArrayResource> generateFromUploaded(
        @PathVariable String uploadId,
        @RequestBody CommentsRequest request
    ) throws Exception {
        WordCommentService.GeneratedDocx generated = wordCommentService.generateFromUpload(uploadId, request.comments());
        return download(generated);
    }

    @PostMapping(value = "/comments/generate", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<ByteArrayResource> generateFromMultipart(
        @RequestParam("file") MultipartFile file,
        @RequestParam("comments") String commentsJson
    ) throws Exception {
        List<CommentDto> comments = objectMapper.readValue(commentsJson, new TypeReference<>() {});
        WordCommentService.GeneratedDocx generated = wordCommentService.generateFromFile(file, comments);
        return download(generated);
    }

    @GetMapping("/health")
    public Map<String, Object> health() {
        return Map.of("ok", true, "backend", "java-aspose");
    }

    private ResponseEntity<ByteArrayResource> download(WordCommentService.GeneratedDocx generated) throws Exception {
        ByteArrayResource resource = new ByteArrayResource(java.nio.file.Files.readAllBytes(generated.path()));
        return ResponseEntity.ok()
            .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + generated.downloadName() + "\"")
            .header("X-Comments-Applied", String.valueOf(generated.applied()))
            .header("X-Comments-Skipped", String.valueOf(generated.skipped()))
            .body(resource);
    }

    @ExceptionHandler(Exception.class)
    public ResponseEntity<Map<String, String>> handleException(Exception error) {
        return ResponseEntity.badRequest().body(Map.of("message", error.getMessage()));
    }

    public record CommentsRequest(List<CommentDto> comments) {
    }
}
