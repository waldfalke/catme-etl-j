package com.catmepim.converter.highvolume.core.writers;

import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.fasterxml.jackson.core.JsonEncoding;
import com.fasterxml.jackson.core.JsonFactory;
import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.util.DefaultPrettyPrinter;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/**
 * Implements {@link IDataWriter} for writing data as a single JSON array.
 * Uses Jackson streaming API for generating JSON output.
 *
 * This class adheres to the contract defined in:
 * java_tools/high-volume-excel-converter/src/main/java/com/catmepim/converter/highvolume/core/writers/IDataWriter-Contract-v1.0.0.md
 *
 * @inv If open, jsonGenerator is not null.
 * @inv rowsWrittenCount accurately reflects data rows written.
 * @inv The output is a single, well-formed JSON array upon successful completion.
 */
public class JsonDataWriter implements IDataWriter {

    private static final Logger logger = LoggerFactory.getLogger(JsonDataWriter.class);
    private static final int BUFFER_SIZE = 128 * 1024; // 128KB buffer for optimal I/O performance as per contract recommendations

    private final ConverterConfig config;
    private Map<Integer, String> headerMap; // 0-indexed column to header name
    private List<String> orderedHeaders;    // Header names in their actual column order

    private OutputStream outputStream;
    private JsonGenerator jsonGenerator;
    private long rowsWrittenCount = 0;
    private boolean firstElementWritten = false;
    private long bytesSinceLastFlush = 0;
    private static final long BYTES_BEFORE_FLUSH = 5 * 1024 * 1024; // ~5MB (~5,000 rows) before forced flush as recommended by contract

    /**
     * Constructs a JsonDataWriter.
     *
     * @param config The converter configuration.
     * @param headerMap The initial header map (column index to header name) provided by ExcelRowListener.
     *                  This map is crucial for determining JSON object keys.
     * @pre config is not null and validated. config.format is JSON.
     * @pre config.outputPath is writable.
     * @pre headerMap is not null (can be empty, but JSON objects will then use column indices as keys).
     * @post The writer is initialized. It is not yet open.
     */
    public JsonDataWriter(ConverterConfig config, Map<Integer, String> headerMap) {
        this.config = config;
        this.headerMap = new TreeMap<>(headerMap); // Ensure consistent key order if iterating over it
        // orderedHeaders will be populated by writeHeader
        logger.info("JsonDataWriter initialized. Output path: {}", config.outputPath);
        if (headerMap.isEmpty()) {
            logger.warn("Initial headerMap is empty. JSON objects will use 0-indexed column indices as keys unless writeHeader provides names.");
        }
    }

    /**
     * @pre config.outputPath is set and the parent directory is writable.
     *      If config.overwrite is false, config.outputPath must not exist.
     * @post The output file is created/truncated, and the JSON generator is initialized, 
     *       writing the opening array bracket '['.
     */
    @Override
    public void open() throws IOException {
        if (jsonGenerator != null) {
            logger.warn("JsonDataWriter is already open. Ignoring open() call.");
            return;
        }

        Path outputPathFile = config.outputPath;
        if (Files.exists(outputPathFile) && !config.overwrite) {
            throw new IOException("Output file " + outputPathFile + " already exists and overwrite is false.");
        }
        Files.createDirectories(outputPathFile.getParent());
        
        // Use BufferedOutputStream with large buffer for improved performance
        outputStream = new BufferedOutputStream(new FileOutputStream(outputPathFile.toFile()), BUFFER_SIZE);
        JsonFactory factory = new JsonFactory();
        
        // Configure Jackson to use lower memory
        factory.configure(JsonFactory.Feature.CANONICALIZE_FIELD_NAMES, false);
        factory.configure(JsonFactory.Feature.INTERN_FIELD_NAMES, false);
        
        jsonGenerator = factory.createGenerator(outputStream, JsonEncoding.UTF8);
        
        // Configure JsonGenerator for memory efficiency
        jsonGenerator.configure(JsonGenerator.Feature.AUTO_CLOSE_TARGET, false);
        jsonGenerator.configure(JsonGenerator.Feature.FLUSH_PASSED_TO_STREAM, true);
        
        // Only use pretty printing if explicitly requested (uses more memory)
        if (config.prettyPrint) {
            jsonGenerator.useDefaultPrettyPrinter();
        }

        jsonGenerator.writeStartArray();
        firstElementWritten = false;
        bytesSinceLastFlush = 0;
        logger.info("JsonDataWriter opened. Output file: {}. Started JSON array.", outputPathFile);
        
        // Force first flush to get the opening bracket written
        jsonGenerator.flush();
    }

    /**
     * @pre The writer is open. headerData is not null.
     * @post Header information is stored for use as keys in JSON objects.
     */
    @Override
    public void writeHeader(Map<Integer, String> headerData) throws IOException {
        if (jsonGenerator == null) throw new IOException("Writer is not open. Call open() first.");
        
        // Ensure headers are sorted by column index for consistent key mapping
        Map<Integer, String> sortedHeaders = new TreeMap<>(headerData);
        this.headerMap = sortedHeaders; // Update internal map
        this.orderedHeaders = new ArrayList<>(sortedHeaders.values());
        logger.debug("JSON headers updated/stored: {}", this.orderedHeaders);
        if (this.orderedHeaders.isEmpty()) {
             logger.warn("Header data is empty. Subsequent JSON objects will use 0-indexed column numbers as keys.");
        }
    }

    /**
     * @pre The writer is open. rowData is not null.
     * @post Row data is written as a JSON object within the main JSON array.
     */
    @Override
    public void writeRow(Map<Integer, String> rowData) throws IOException {
        if (jsonGenerator == null) throw new IOException("Writer is not open. Call open() first.");

        jsonGenerator.writeStartObject();
        Map<Integer, String> sortedRowData = new TreeMap<>(rowData); // Ensure data is processed in column order

        for (Map.Entry<Integer, String> entry : sortedRowData.entrySet()) {
            Integer colIndex = entry.getKey();
            String cellValue = entry.getValue();
            String key = (this.orderedHeaders != null && colIndex < this.orderedHeaders.size() && this.orderedHeaders.get(colIndex) != null) 
                         ? this.orderedHeaders.get(colIndex) 
                         : Integer.toString(colIndex);
            if (key.isEmpty()) key = Integer.toString(colIndex); // Fallback for empty header names

            jsonGenerator.writeStringField(key, cellValue);
        }
        jsonGenerator.writeEndObject();
        firstElementWritten = true;
        rowsWrittenCount++;
        
        // Estimate bytes written and force flush periodically
        // This is an approximation - each row might be different
        bytesSinceLastFlush += estimateJsonObjectSize(sortedRowData);
        
        if (bytesSinceLastFlush >= BYTES_BEFORE_FLUSH) {
            logger.debug("Memory threshold reached ({} MB accumulated). Forcing flush to disk.", 
                    String.format("%.2f", bytesSinceLastFlush / (1024.0 * 1024.0)));
            flush();
            bytesSinceLastFlush = 0;
            
            // Help garbage collector by clearing reference caches
            System.gc();
        }
    }

    /**
     * Estimates the size of a JSON object based on row data.
     * This is a rough approximation to determine when to flush.
     * 
     * @param rowData The data map for a row
     * @return Estimated byte size of the JSON representation
     */
    private long estimateJsonObjectSize(Map<Integer, String> rowData) {
        // Base size for {}, commas, etc.
        long estimatedSize = 10;
        
        // Add size for each field (key + value + quotes + colon)
        for (Map.Entry<Integer, String> entry : rowData.entrySet()) {
            String key = entry.getKey().toString();
            String value = entry.getValue();
            
            // key: quotes + key length + colon + space
            estimatedSize += 4 + key.length();
            
            // value: quotes + value length
            estimatedSize += 2 + (value != null ? value.length() : 4);
            
            // comma + newline
            estimatedSize += 2;
        }
        
        return estimatedSize;
    }

    /**
     * @pre The writer is open.
     * @post JsonGenerator buffer is flushed to disk.
     */
    @Override
    public void flush() throws IOException {
        if (jsonGenerator == null) throw new IOException("Writer is not open. Call open() first.");
        logger.debug("Flushing JsonDataWriter to disk for file: {}", config.outputPath);
        jsonGenerator.flush();
        outputStream.flush(); // Ensure data actually goes to disk
    }

    /**
     * @pre Writer may be open or closed.
     * @post JSON array is closed, and all resources are flushed and closed.
     */
    @Override
    public void close() throws IOException {
        logger.info("Closing JsonDataWriter. Total rows written: {}. Output file: {}", rowsWrittenCount, config.outputPath);
        if (jsonGenerator != null) {
            try {
                if (!jsonGenerator.isClosed()) {
                    jsonGenerator.writeEndArray(); // Close the main JSON array
                    jsonGenerator.flush(); // Ensure the end array is written
                    jsonGenerator.close(); // This also flushes the outputStream
                }
            } catch (IOException e) {
                logger.error("Error closing JsonGenerator: {}", e.getMessage(), e);
                throw e;
            } finally {
                // Explicitly close outputStream in case JsonGenerator didn't
                if (outputStream != null) {
                    try {
                        outputStream.close();
                    } catch (IOException e) {
                        logger.warn("Error when closing output stream: {}", e.getMessage());
                    }
                }
                jsonGenerator = null;
                outputStream = null;
            }
        }
    }

    /**
     * @pre Writer can be in any state.
     * @post Returns total data rows written.
     */
    @Override
    public long getRowsWrittenCount() {
        return rowsWrittenCount;
    }
} 