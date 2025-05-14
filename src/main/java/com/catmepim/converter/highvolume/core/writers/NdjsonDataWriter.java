package com.catmepim.converter.highvolume.core.writers;

import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.fasterxml.jackson.core.JsonEncoding;
import com.fasterxml.jackson.core.JsonFactory;
import com.fasterxml.jackson.core.JsonGenerator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/**
 * Implements IDataWriter to write data in NDJSON (Newline Delimited JSON) format.
 * Each row is written as a separate JSON object on a new line.
 * Adheres to contract: NdjsonDataWriter-Contract-v1.0.0.md
 */
public class NdjsonDataWriter implements IDataWriter {
    private static final Logger LOG = LoggerFactory.getLogger(NdjsonDataWriter.class);

    private final ConverterConfig config;
    private BufferedWriter writer;
    private JsonGenerator jsonGenerator;
    private final JsonFactory jsonFactory;
    private List<String> orderedHeaders;
    private long rowsWrittenCount;
    private final Map<Integer, String> initialHeaderMap; // Keep initial map for reference if needed

    /**
     * @pre {@code config != null}
     * @pre {@code config.outputPath != null}
     * @pre {@code initialHeaderMap != null}
     * @post A new NdjsonDataWriter instance is created.
     */
    public NdjsonDataWriter(ConverterConfig config, Map<Integer, String> initialHeaderMap) {
        if (config == null) {
            throw new IllegalArgumentException("ConverterConfig cannot be null.");
        }
        if (config.outputPath == null) {
            throw new IllegalArgumentException("Output path cannot be null for NDJSON writer.");
        }
        if (initialHeaderMap == null) {
            throw new IllegalArgumentException("Initial header map cannot be null.");
        }
        this.config = config;
        this.initialHeaderMap = initialHeaderMap; // Store if needed, or directly use for orderedHeaders
        this.orderedHeaders = new ArrayList<>();
        this.rowsWrittenCount = 0;
        this.jsonFactory = new JsonFactory();
        // jsonGenerator and writer are initialized in open()
    }

    /**
     * @pre {@code this.config.outputPath} is valid.
     * @post Output file is opened/created. `jsonGenerator` and `writer` are initialized.
     * @throws IOException if an I/O error occurs.
     */
    @Override
    public void open() throws IOException {
        Path outputFile = config.outputPath;
        LOG.info("Opening NDJSON writer for file: {}", outputFile.toAbsolutePath());

        if (Files.exists(outputFile) && !config.overwrite) {
            String errorMessage = "Output file " + outputFile.toAbsolutePath() + " already exists and overwrite is false.";
            LOG.error(errorMessage);
            throw new IOException(errorMessage);
        }

        Path parentDir = outputFile.getParent();
        if (parentDir != null && !Files.exists(parentDir)) {
            Files.createDirectories(parentDir);
            LOG.info("Created parent directory: {}", parentDir.toAbsolutePath());
        }

        FileOutputStream fos = new FileOutputStream(outputFile.toFile()); // Overwrites by default if file exists
        this.writer = new BufferedWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));
        this.jsonGenerator = jsonFactory.createGenerator(this.writer);
        // No pretty printing for NDJSON, each object is on one line.

        LOG.info("NDJSON writer opened successfully for: {}", outputFile.toAbsolutePath());
    }

    /**
     * @pre Writer is open. `headerData != null`.
     * @post `orderedHeaders` is populated.
     */
    @Override
    public void writeHeader(Map<Integer, String> headerData) throws IOException {
        if (headerData == null) {
            LOG.warn("Header data is null. Headers will be empty.");
            this.orderedHeaders = new ArrayList<>();
            return;
        }
        // Use TreeMap to sort by column index, then extract values to maintain order
        Map<Integer, String> sortedHeaders = new TreeMap<>(headerData);
        this.orderedHeaders = new ArrayList<>(sortedHeaders.values());
        LOG.debug("Stored ordered headers: {}", this.orderedHeaders);
        // No actual header writing to file for NDJSON, headers are used per row object.
    }

    /**
     * @pre Writer is open. `rowData != null`.
     * @post Row is written as a JSON object followed by a newline. `rowsWrittenCount` incremented.
     * @throws IOException if an I/O error occurs.
     */
    @Override
    public void writeRow(Map<Integer, String> rowData) throws IOException {
        if (rowData == null) {
            LOG.warn("Skipping null row data.");
            return;
        }
        if (jsonGenerator == null) {
            throw new IOException("JsonGenerator is not initialized. Writer might not be open.");
        }

        jsonGenerator.writeStartObject();
        // Ensure data is written in header order if headers are available
        Map<Integer, String> sortedRowData = new TreeMap<>(rowData); // Sort row data by column index

        for (Map.Entry<Integer, String> entry : sortedRowData.entrySet()) {
            Integer colIndex = entry.getKey();
            String cellValue = entry.getValue();
            String headerName;

            if (colIndex < orderedHeaders.size()) {
                headerName = orderedHeaders.get(colIndex);
            } else {
                // Fallback if row has more columns than headers, or headers are missing
                headerName = Integer.toString(colIndex);
                LOG.trace("Missing header for column index {}, using index as key.", colIndex);
            }
            jsonGenerator.writeStringField(headerName, cellValue);
        }
        jsonGenerator.writeEndObject();
        jsonGenerator.flush(); // Flush the JsonGenerator's internal buffer to the writer
        writer.newLine();      // Write the newline character using the BufferedWriter
        writer.flush();        // Flush the BufferedWriter to ensure the line is written to the file

        rowsWrittenCount++;
    }

    /**
     * @pre Writer is open.
     * @post Buffered data is flushed.
     * @throws IOException if an I/O error occurs.
     */
    @Override
    public void flush() throws IOException {
        if (jsonGenerator != null) {
            jsonGenerator.flush();
        }
        if (writer != null) {
            writer.flush();
        }
        LOG.debug("NDJSON writer flushed.");
    }

    /**
     * @pre Writer can be open or closed.
     * @post Resources are closed. Writer is no longer usable.
     * @throws IOException if an I/O error occurs.
     */
    @Override
    public void close() throws IOException {
        LOG.info("Closing NDJSON writer. Total rows written: {}", rowsWrittenCount);
        try {
            if (jsonGenerator != null) {
                // Don't close the underlying writer here, as JsonGenerator might not own it
                // or might close it in a way that conflicts with BufferedWriter
                // jsonGenerator.close(); //This will also close the writer, which we manage separately
                jsonGenerator.flush(); // Ensure anything in JsonGenerator buffer is written
            }
        } finally {
            // Close the BufferedWriter, which will in turn close the FileOutputStream
            if (writer != null) {
                try {
                    writer.close(); // This will flush and then close
                } catch (IOException e) {
                    LOG.error("Error closing BufferedWriter for NDJSON: {}", e.getMessage(), e);
                    // Rethrow or suppress, depending on policy. For now, rethrow to signal issue.
                    throw e;
                }
            }
        }
        LOG.info("NDJSON writer closed successfully.");
    }

    /**
     * @return Total rows successfully written.
     */
    @Override
    public long getRowsWrittenCount() {
        return rowsWrittenCount;
    }
} 