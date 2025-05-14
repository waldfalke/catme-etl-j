package com.catmepim.converter.highvolume.core;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.catmepim.converter.highvolume.core.writers.CsvDataWriter;
import com.catmepim.converter.highvolume.core.writers.IDataWriter;
import com.catmepim.converter.highvolume.core.writers.JsonDataWriter;
import com.catmepim.converter.highvolume.core.writers.NdjsonDataWriter;

/**
 * An EasyExcel {@link AnalysisEventListener} that processes rows from an Excel sheet.
 * Handles header identification, data reception, and delegates writing to
 * format-specific writers based on the provided {@link ConverterConfig}.
 *
 * Adheres to contract: ExcelRowListener-Contract-v1.0.0.md
 *
 * @invariant The internal rowCount accurately reflects the number of data rows processed via invoke().
 * @invariant The ConverterConfig instance (config) is not modified by this listener.
 * @invariant If a data writer is active, it corresponds to the format specified in config.format.
 * @contract ExcelRowListener-Contract-v1.0.0.md, CatmePoiSheetContentsHandler-Contract-v1.0.0.md
 */
public class ExcelRowListener extends AnalysisEventListener<Map<Integer, String>> {

    private static final Logger logger = LoggerFactory.getLogger(ExcelRowListener.class);
    private static final int LOG_INTERVAL = 10000; // Log progress every 10,000 rows

    private final ConverterConfig config;
    private Map<Integer, String> headerMap = new LinkedHashMap<>();
    private long rowCount = 0;
    private long rowsSinceLastLog;
    private long totalRowsSuccessfullyWritten;
    private IDataWriter dataWriter;
    private long lastLogTimeMillis = System.currentTimeMillis();

    /**
     * Constructs an ExcelRowListener.
     *
     * @param config The converter configuration, which dictates output format, paths, etc.
     * @pre config is not null and has been validated externally.
     * @pre config.format specifies a valid OutputFormat.
     * @pre config.outputPath (for JSON/NDJSON) or config.tempDir (for CSV) are writable.
     * @pre config.batchSize > 0.
     * @pre config.headerRow >= 0.
     * @post The listener is initialized with the configuration.
     * @post The appropriate data writer is instantiated, opened, and ready for writing.
     */
    public ExcelRowListener(ConverterConfig config) {
        this.config = config;
        this.rowsSinceLastLog = 0;
        this.totalRowsSuccessfullyWritten = 0;

        try {
            if (config.format == ConverterConfig.OutputFormat.CSV) {
                Path tempDirPath = config.tempDir;
                if (!Files.exists(tempDirPath)) {
                    Files.createDirectories(tempDirPath);
                    logger.info("Created temporary directory: {}", tempDirPath.toAbsolutePath());
                }
            }

            switch (config.format) {
                case CSV:
                    this.dataWriter = new CsvDataWriter(config, this.headerMap);
                    break;
                case JSON:
                    this.dataWriter = new JsonDataWriter(config, this.headerMap);
                    break;
                case NDJSON:
                    this.dataWriter = new NdjsonDataWriter(config, this.headerMap);
                    break;
                default:
                    String errorMsg = "Unsupported output format in ExcelRowListener: " + config.format;
                    logger.error(errorMsg);
                    throw new IllegalArgumentException(errorMsg);
            }
            this.dataWriter.open();
        } catch (IOException e) {
            logger.error("Failed to open IDataWriter for format {}: {}", config.format, e.getMessage(), e);
            throw new RuntimeException("Failed to initialize or open data writer", e);
        } catch (UnsupportedOperationException | IllegalArgumentException e) {
            logger.error("Failed to determine or initialize data writer: {}", e.getMessage());
            throw e;
        }
    }

    /**
     * Called by EasyExcel or POI handler when the header row is encountered.
     *
     * @param headMap A map where keys are 0-indexed column numbers and values are header cell content.
     * @param context The analysis context from EasyExcel, or null if called from POI handler.
     * @pre headMap contains data from the row identified by EasyExcel as the header row, or POI handler.
     * @pre context may be null (if called from POI handler).
     * @post Header data is stored internally and written to the output via the dataWriter.
     * @throws RuntimeException if writing header fails.
     * @contract ExcelRowListener-Contract-v1.0.0.md, CatmePoiSheetContentsHandler-Contract-v1.0.0.md
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        String sheetName = (context != null && context.readSheetHolder() != null)
            ? context.readSheetHolder().getSheetName()
            : "<POI/Unknown>";
        if (context == null) {
            logger.debug("invokeHeadMap called with null context (likely from POI handler). Using sheetName='{}'.", sheetName);
        }
        logger.info("Header row processed. Sheet: '{}', Headers: {}", sheetName, headMap);
        this.headerMap.clear();
        this.headerMap.putAll(headMap);
        if (dataWriter != null) {
            try {
                dataWriter.writeHeader(this.headerMap);
            } catch (IOException e) {
                logger.error("Error writing header: {}", e.getMessage(), e);
                throw new RuntimeException("Error writing header data", e);
            }
        }
    }

    /**
     * Called by EasyExcel or POI handler for each data row (after the header row).
     *
     * @param data A map where keys are 0-indexed column numbers and values are cell content for the current row.
     * @param context The analysis context from EasyExcel, or null if called from POI handler.
     * @pre data contains cell values for a single data row.
     * @pre context may be null (if called from POI handler).
     * @post The row data is passed to the dataWriter.
     * @post rowCount is incremented.
     * @post CSV chunking logic is triggered by the dataWriter if applicable.
     * @throws RuntimeException if writing row fails and continueOnError is false.
     * @contract ExcelRowListener-Contract-v1.0.0.md, CatmePoiSheetContentsHandler-Contract-v1.0.0.md
     */
    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {
        String sheetName = (context != null && context.readSheetHolder() != null)
            ? context.readSheetHolder().getSheetName()
            : "<POI/Unknown>";
        if (context == null) {
            logger.trace("invoke called with null context (likely from POI handler). Using sheetName='{}'.", sheetName);
        }
        logger.trace("Processing data row {}: {} (Sheet: '{}')", rowCount, data, sheetName);
        rowCount++;

        if (dataWriter != null) {
            try {
                dataWriter.writeRow(data);
                totalRowsSuccessfullyWritten++;
                rowsSinceLastLog++;
                if (rowsSinceLastLog >= LOG_INTERVAL) {
                    long now = System.currentTimeMillis();
                    long intervalMillis = now - lastLogTimeMillis;
                    double speed = intervalMillis > 0 ? (rowsSinceLastLog * 1000.0 / intervalMillis) : 0.0;
                    logger.info("Processed {} rows so far... ({} rows in {} ms, ~{}/sec)",
                        totalRowsSuccessfullyWritten, rowsSinceLastLog, intervalMillis, String.format("%.2f", speed));
                    lastLogTimeMillis = now;
                    rowsSinceLastLog = 0;
                }
            } catch (IOException e) {
                logger.error("Error writing data row {}: {}", rowCount, e.getMessage(), e);
                if (!config.continueOnError) {
                    throw new RuntimeException("Error writing data row " + rowCount, e);
                }
            }
        }
        
        if (rowCount % config.batchSize == 0) {
            logger.info("Listener has processed {} data rows. Flushing writer.", rowCount);
            if (dataWriter != null) {
                try {
                    dataWriter.flush();
                } catch (IOException e) {
                    logger.warn("Error flushing data writer: {}", e.getMessage(), e);
                }
            }
        }
    }

    /**
     * Called by EasyExcel or POI handler after all rows in a sheet have been processed.
     *
     * @param context The analysis context from EasyExcel, or null if called from POI handler.
     * @pre context may be null (if called from POI handler).
     * @pre All rows from the sheet have been passed to invoke() or invokeHeadMap().
     * @post All open data writers are flushed and closed via the dataWriter.
     * @post A summary including rows written by dataWriter is logged.
     * @throws RuntimeException if closing data writer fails.
     * @contract ExcelRowListener-Contract-v1.0.0.md, CatmePoiSheetContentsHandler-Contract-v1.0.0.md
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        String sheetName = (context != null && context.readSheetHolder() != null)
            ? context.readSheetHolder().getSheetName()
            : "<POI/Unknown>";
        if (context == null) {
            logger.debug("doAfterAllAnalysed called with null context (likely from POI handler). Using sheetName='{}'.", sheetName);
        }
        logger.info("Finished analysing sheet '{}'. Total data rows processed by listener: {}.", sheetName, rowCount);
        if (dataWriter != null) {
            try {
                dataWriter.close();
                logger.info("IDataWriter closed. Total rows successfully written: {}", dataWriter.getRowsWrittenCount());
            } catch (IOException e) {
                logger.error("Error closing data writer: {}", e.getMessage(), e);
                throw new RuntimeException("Error closing data writer", e);
            }
        } else {
             logger.warn("No data writer was initialized, so no data was written.");
        }
    }

    /**
     * Explicitly flushes the data writer to ensure all pending data is written.
     * Used for checkpoint operations to ensure data is persisted even during long-running processes.
     *
     * @throws IOException if an I/O error occurs during flushing
     * @post Any buffered data in the writer is flushed to the output
     */
    public void flushWriter() throws IOException {
        if (dataWriter != null) {
            dataWriter.flush();
            logger.info("Data writer explicitly flushed for checkpoint. Rows processed so far: {}", rowCount);
        } else {
            logger.warn("Attempted to flush data writer but none is initialized.");
        }
    }

    /**
     * Called by EasyExcel if an exception occurs during a read operation.
     *
     * @param exception The exception that occurred.
     * @param context The analysis context from EasyExcel, or null if called from POI handler.
     * @pre exception is not null.
     * @pre context may be null (if called from POI handler).
     * @post Processing may halt or continue based on config.continueOnError.
     * @throws Exception always rethrows if continueOnError is false.
     * @contract ExcelRowListener-Contract-v1.0.0.md, CatmePoiSheetContentsHandler-Contract-v1.0.0.md
     */
    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        String sheetName = (context != null && context.readSheetHolder() != null)
            ? context.readSheetHolder().getSheetName()
            : "<POI/Unknown>";
        if (context == null) {
            logger.warn("onException called with null context (likely from POI handler). Using sheetName='{}'.", sheetName);
        }
        logger.error("Exception during Excel read operation at/near row index {}: {}. Sheet: '{}'",
                     (context != null && context.readRowHolder() != null ? context.readRowHolder().getRowIndex() : "N/A"),
                     exception.getMessage(),
                     sheetName,
                     exception);
        if (dataWriter != null && !config.continueOnError) {
            try {
                logger.info("Attempting to close data writer due to unrecoverable exception...");
                dataWriter.close();
            } catch (IOException ioe) {
                logger.error("Additionally, error closing data writer during exception handling: {}", ioe.getMessage(), ioe);
            }
        }
        if (!config.continueOnError) {
            logger.error("Halting processing due to error and continueOnError=false.");
            throw exception;
        } else {
            logger.warn("continueOnError=true. Attempting to continue processing despite the error. Current row data may be lost.");
        }
    }

    /**
     * Gets the total number of rows successfully written by the underlying data writer.
     * @return Number of rows written, or 0 if no writer was initialized.
     */
    public long getTotalRowsSuccessfullyWritten() {
        if (dataWriter != null) {
            return dataWriter.getRowsWrittenCount();
        }
        return 0;
    }
}