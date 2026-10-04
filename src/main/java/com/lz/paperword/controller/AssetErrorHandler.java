package com.lz.paperword.controller;

import com.lz.paperword.core.docx.ImageAssetException;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import java.util.Map;

/** Shared contract for Word and PDF routes: a bad image is never silently dropped. */
@RestControllerAdvice
public final class AssetErrorHandler {
    @ExceptionHandler(ImageAssetException.class)
    public ResponseEntity<Map<String,String>> imageFailure(ImageAssetException e) {
        HttpStatus status=switch(e.getCode()) {
            case "IMAGE_ASSET_TOO_LARGE", "IMAGE_ASSET_TOO_MANY_PIXELS" -> HttpStatus.PAYLOAD_TOO_LARGE;
            case "IMAGE_ASSET_DISABLED", "IMAGE_ASSET_CONFIG" -> HttpStatus.SERVICE_UNAVAILABLE;
            default -> HttpStatus.BAD_REQUEST;
        };
        return ResponseEntity.status(status).body(Map.of("code",e.getCode(),"message",e.getMessage()));
    }
}
