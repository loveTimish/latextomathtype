package com.lz.paperword.controller;

import com.lz.paperword.model.ExportRequestValidationException;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import java.util.Map;

@RestControllerAdvice
public final class ExportRequestErrorHandler {
    @ExceptionHandler(ExportRequestValidationException.class)
    public ResponseEntity<Map<String,String>> invalidLayout(ExportRequestValidationException error) {
        return ResponseEntity.badRequest().body(Map.of("code","INVALID_EXPORT_REQUEST","message",error.getMessage()));
    }
}
