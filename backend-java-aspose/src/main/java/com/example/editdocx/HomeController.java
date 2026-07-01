package com.example.editdocx;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.Map;

@RestController
public class HomeController {

    @GetMapping("/")
    public Map<String, Object> home() {
        return Map.of(
            "service", "editDocx Java Aspose API",
            "status", "ok",
            "health", "/api/health",
            "docs", "Upload DOCX via POST /api/docx/upload"
        );
    }
}
