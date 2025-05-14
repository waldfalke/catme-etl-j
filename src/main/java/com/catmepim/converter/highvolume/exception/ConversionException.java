package com.catmepim.converter.highvolume.exception;

/**
 * Exception thrown when a conversion error occurs in the HighVolumeExcelConverter.
 * <p>
 * This exception is used to indicate problems during the Excel conversion process,
 * such as file access issues, invalid data, or format-specific errors.
 * <p>
 * Contract reference: HighVolumeExcelConverter-Contract-v2.0.1.md (section on error handling)
 */
public class ConversionException extends RuntimeException {

    /**
     * Constructs a new ConversionException with the specified detail message.
     *
     * @param message the detail message
     */
    public ConversionException(String message) {
        super(message);
    }

    /**
     * Constructs a new ConversionException with the specified detail message and cause.
     *
     * @param message the detail message
     * @param cause the cause
     */
    public ConversionException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs a new ConversionException with the specified cause.
     *
     * @param cause the cause
     */
    public ConversionException(Throwable cause) {
        super(cause);
    }
} 