package com.lz.paperword.core.docx;

/** A failed requested image aborts the export instead of silently omitting content. */
public final class ImageAssetException extends RuntimeException {
    private final String code;

    public ImageAssetException(String code, String message) {
        super(message);
        this.code = code;
    }

    public ImageAssetException(String code, String message, Throwable cause) {
        super(message, cause);
        this.code = code;
    }

    public String getCode() {
        return code;
    }
}
