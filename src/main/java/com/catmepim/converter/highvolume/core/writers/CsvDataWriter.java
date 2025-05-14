package com.catmepim.converter.highvolume.core.writers;

import com.catmepim.converter.highvolume.config.ConverterConfig;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/**
 * Implements {@link IDataWriter} for writing data to CSV files.
 * Handles CSV chunking based on configuration.
 *
 * This class adheres to the contract defined in:
 * java_tools/high-volume-excel-converter/src/main/java/com/catmepim/converter/highvolume/core/writers/IDataWriter-Contract-v1.0.0.md
 *
 * @inv If open, csvPrinter and writer are not null.
 * @inv rowsWrittenCount accurately reflects data rows written across all chunks.
 * @inv currentChunkIndex is non-negative.
 */
public class CsvDataWriter implements IDataWriter {

    private static final Logger logger = LoggerFactory.getLogger(CsvDataWriter.class);

    private final ConverterConfig config;
    private final Map<Integer, String> initialHeaderMap; // Headers received from ExcelRowListener
    private List<String> orderedHeaders; // Headers in the order they should be written

    private BufferedWriter writer;
    private CSVPrinter csvPrinter;

    private int currentChunkIndex = 0;
    private long rowsInCurrentChunk = 0;
    private long totalRowsWrittenCount = 0;
    private boolean headerWrittenForCurrentChunk = false;
    private Path currentChunkFilePath;

    /**
     * Constructs a CsvDataWriter.
     *
     * @param config The converter configuration.
     * @param initialHeaderMap The header map (column index to header name) from ExcelRowListener. Can be empty.
     * @pre config is not null and validated. config.format is CSV.
     * @pre config.tempDir is writable. config.inputFile is set.
     * @pre initialHeaderMap is not null.
     * @post The writer is initialized with config and headers. It is not yet open.
     */
    public CsvDataWriter(ConverterConfig config, Map<Integer, String> initialHeaderMap) {
        this.config = config;
        this.initialHeaderMap = new TreeMap<>(initialHeaderMap); // Ensure consistent order if used directly
        // orderedHeaders will be set in writeHeader or derived if initialHeaderMap is empty but writeHeader is called
        logger.info("CsvDataWriter initialized. Temp directory: {}, Batch size: {}", config.tempDir, config.batchSize);
    }

    /**
     * @pre config.tempDir is writable.
     * @post The first CSV chunk file is created and open for writing. Headers are written if available.
     */
    @Override
    public void open() throws IOException {
        if (csvPrinter != null) {
            logger.warn("CsvDataWriter is already open. Ignoring open() call.");
            return;
        }
        Files.createDirectories(config.tempDir);
        openNewChunk(); // Opens chunk 0 or 1 (will be 1 after this call)
        logger.info("CsvDataWriter opened. Initial chunk: {}", currentChunkFilePath);
    }

    private void openNewChunk() throws IOException {
        closeCurrentChunkResources(); // Close previous if any

        currentChunkIndex++;
        rowsInCurrentChunk = 0;
        headerWrittenForCurrentChunk = false;

        String originalFileName = Paths.get(config.inputFile.getFileName().toString()).getFileName().toString();
        String baseName = originalFileName.contains(".") ? originalFileName.substring(0, originalFileName.lastIndexOf('.')) : originalFileName;
        String chunkFileName = String.format("%s-chunk-%d.csv", baseName, currentChunkIndex);
        currentChunkFilePath = config.tempDir.resolve(chunkFileName);

        logger.info("Opening new CSV chunk: {}", currentChunkFilePath);
        writer = Files.newBufferedWriter(currentChunkFilePath, StandardOpenOption.CREATE, StandardOpenOption.WRITE, StandardOpenOption.TRUNCATE_EXISTING);
        // Consider using config.csvDelimiter if we add it
        csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT);

        // If headers are known and need to be written to this new chunk
        if (this.orderedHeaders != null && !this.orderedHeaders.isEmpty()) {
            logger.debug("Writing headers to new chunk: {}", this.orderedHeaders);
            csvPrinter.printRecord(this.orderedHeaders);
            headerWrittenForCurrentChunk = true;
        }
    }

    private void closeCurrentChunkResources() throws IOException {
        if (csvPrinter != null) {
            try {
                csvPrinter.flush();
                csvPrinter.close(true); // Close the CSVPrinter and its underlying writer (if closeTarget=true)
            } catch (IOException e) {
                logger.error("Error closing CSVPrinter for chunk {}: {}", currentChunkFilePath, e.getMessage(), e);
                throw e; // Rethrow as this is a critical operation
            } finally {
                csvPrinter = null;
                writer = null; // writer is closed by CSVPrinter normally
                logger.info("Closed CSV chunk: {}", currentChunkFilePath);
            }
        }
    }

    /**
     * @pre The writer is open (csvPrinter is not null). headerData is not null.
     * @post Header information is stored and written to the current CSV chunk if not already done.
     */
    @Override
    public void writeHeader(Map<Integer, String> headerData) throws IOException {
        if (csvPrinter == null) throw new IOException("Writer is not open. Call open() first.");
        
        // Store and order headers. Use TreeMap for consistent ordering by column index.
        Map<Integer, String> sortedHeaderData = new TreeMap<>(headerData);
        this.orderedHeaders = new ArrayList<>(sortedHeaderData.values());

        if (!headerWrittenForCurrentChunk && !this.orderedHeaders.isEmpty()) {
            logger.debug("Writing explicit headers to current chunk {}: {}", currentChunkFilePath, this.orderedHeaders);
            csvPrinter.printRecord(this.orderedHeaders);
            headerWrittenForCurrentChunk = true;
        }
    }

    /**
     * @pre The writer is open. rowData is not null.
     * @post Row data is written. If batch size reached, current chunk is closed and new one opened.
     */
    @Override
    public void writeRow(Map<Integer, String> rowData) throws IOException {
        if (csvPrinter == null) throw new IOException("Writer is not open. Call open() first.");

        if (rowsInCurrentChunk >= config.batchSize) {
            logger.info("Batch size ({}) reached for chunk {}. Closing and opening next chunk.", config.batchSize, currentChunkFilePath);
            openNewChunk();
        }

        List<String> printableRow = new ArrayList<>();
        if (this.orderedHeaders != null && !this.orderedHeaders.isEmpty()) {
            // If headers are defined, ensure row data aligns with header order
            // This assumes rowData keys (column indices) match what initialHeaderMap provided
            // For simplicity, let's assume rowData is a TreeMap from ExcelRowListener or is already sorted by key
            // Or, more robustly, iterate based on initialHeaderMap's keys if they define the column order
            Map<Integer, String> sortedRowData = new TreeMap<>(rowData);
            printableRow.addAll(sortedRowData.values());
        } else {
            // No headers defined yet, write data as received (ordered by column index)
            Map<Integer, String> sortedRowData = new TreeMap<>(rowData);
            printableRow.addAll(sortedRowData.values());
        }
        
        csvPrinter.printRecord(printableRow);
        rowsInCurrentChunk++;
        totalRowsWrittenCount++;
    }

    /**
     * @pre The writer is open.
     * @post CSVPrinter buffer is flushed.
     */
    @Override
    public void flush() throws IOException {
        if (csvPrinter == null) throw new IOException("Writer is not open. Call open() first.");
        logger.debug("Flushing CsvDataWriter for chunk: {}", currentChunkFilePath);
        csvPrinter.flush();
    }

    /**
     * @pre Writer may be open or closed.
     * @post All resources are closed. Final data is flushed.
     */
    @Override
    public void close() throws IOException {
        logger.info("Closing CsvDataWriter. Total rows written: {}. Last chunk: {}", totalRowsWrittenCount, currentChunkFilePath);
        closeCurrentChunkResources();
        // Final sanity check in case open() was never called but close() is.
        if (csvPrinter != null) {
             csvPrinter.close(true);
             csvPrinter = null;
        }
        if (writer != null) { // Should be closed by CSVPrinter, but as a safeguard
            writer.close();
            writer = null;
        }
    }

    /**
     * @pre Writer can be in any state.
     * @post Returns total data rows written across all chunks.
     */
    @Override
    public long getRowsWrittenCount() {
        return totalRowsWrittenCount;
    }
} 