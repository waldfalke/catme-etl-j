package com.catmepim.converter.highvolume.core.writers;

import java.io.IOException;
import java.util.Map;

/**
 * Defines a common interface for writing processed Excel row data to different output formats.
 * Implementations of this interface handle the specifics of each format (e.g., CSV, JSON, NDJSON).
 *
 * This interface and its implementations adhere to the contract defined in:
 * java_tools/high-volume-excel-converter/src/main/java/com/catmepim/converter/highvolume/core/writers/IDataWriter-Contract-v1.0.0.md
 *
 * @inv Implementations must ensure that all resources (files, streams) are properly managed and closed, typically in the close() method.
 * @inv The getRowsWrittenCount() method should accurately reflect the number of data rows successfully processed by writeRow().
 */
public interface IDataWriter extends AutoCloseable { // Extending AutoCloseable for try-with-resources compatibility

    /**
     * Prepares the writer for operation. This may include opening files or streams.
     * This method should be idempotent; calling it multiple times on an already open writer should not cause errors.
     *
     * @throws IOException if an I/O error occurs during opening or preparation.
     * @pre The writer is in a state where it can be opened (e.g., configuration is valid).
     * @post The writer is ready to accept data via writeHeader() and writeRow(). Any necessary files/streams are opened.
     */
    void open() throws IOException;

    /**
     * Writes header data to the output, if applicable for the specific format.
     * For formats like CSV, this would write the header row. For others (like JSON objects), 
     * this might store the header internally to be used as keys for subsequent data rows.
     *
     * @param headerData A map where keys are 0-indexed column indices and values are header strings.
     * @throws IOException if an I/O error occurs during writing.
     * @pre The writer is open.
     * @pre headerData is not null.
     * @post Header information is processed. For formats that write headers, the header is written to the output.
     */
    void writeHeader(Map<Integer, String> headerData) throws IOException;

    /**
     * Writes a single row of data to the output.
     *
     * @param rowData A map where keys are 0-indexed column indices and values are cell content strings.
     * @throws IOException if an I/O error occurs during writing.
     * @pre The writer is open.
     * @pre rowData is not null.
     * @post The provided rowData is written to the output stream/file in the target format.
     *       The internal count for getRowsWrittenCount() is incremented if writing was successful.
     */
    void writeRow(Map<Integer, String> rowData) throws IOException;

    /**
     * Flushes any buffered data to the underlying output stream/file.
     * This is important to ensure data is persisted, especially before closing or during long operations.
     *
     * @throws IOException if an I/O error occurs during flushing.
     * @pre The writer is open.
     * @post All buffered data is written to the destination.
     */
    void flush() throws IOException;

    /**
     * Closes the writer and releases any system resources associated with it (e.g., file streams).
     * This method should typically call flush() before closing to ensure all data is written.
     * This method is also called by the try-with-resources statement if IDataWriter implements AutoCloseable.
     *
     * @throws IOException if an I/O error occurs during closing.
     * @pre The writer may be open or closed. If open, it will attempt to flush and close.
     * @post All resources held by the writer are released. The writer is no longer usable for writing.
     */
    @Override // From AutoCloseable
    void close() throws IOException;

    /**
     * Returns the total number of data rows successfully written by this writer instance.
     * This count should only include rows passed to writeRow() that were successfully processed.
     *
     * @return The count of successfully written data rows.
     * @pre The writer can be in any state (open or closed).
     * @post The returned count reflects the number of rows written. This value does not change after the writer is closed.
     */
    long getRowsWrittenCount();
} 