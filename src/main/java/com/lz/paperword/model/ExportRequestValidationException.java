package com.lz.paperword.model;

/** A bounded document-layout option is invalid; safe to explain to its caller. */
public final class ExportRequestValidationException extends IllegalArgumentException {
    public ExportRequestValidationException(String message) { super(message); }
}
