package com.catmepim.converter.highvolume.exception;

import java.io.IOException;

/**
 * Exception thrown when a potential zip bomb is detected.
 * This triggers the fallback to manual, safer ZIP extraction.
 * 
 * <p>Reference: HighVolumeExcelConverter-Contract-v2.0.1.md (§9)</p>
 */
public class ZipBombDetectedException extends IOException {
    
    private static final long serialVersionUID = 1L;

    /**
     * Constructs a new ZipBombDetectedException with the specified detail message.
     *
     * @param message The detail message
     */
    public ZipBombDetectedException(String message) {
        super(message);
    }

    /**
     * Constructs a new ZipBombDetectedException with the specified detail message and cause.
     *
     * @param message The detail message
     * @param cause The cause
     */
    public ZipBombDetectedException(String message, Throwable cause) {
        super(message, cause);
    }
} 