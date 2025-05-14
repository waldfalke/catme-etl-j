package com.catmepim.converter.highvolume.strategy;

import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.catmepim.converter.highvolume.config.ConverterConfig.OutputFormat;
import com.catmepim.converter.highvolume.core.ExcelRowListener;
import com.catmepim.converter.highvolume.core.writers.CsvDataWriter;
import com.catmepim.converter.highvolume.core.writers.JsonDataWriter;
import com.catmepim.converter.highvolume.core.writers.IDataWriter;
import com.catmepim.converter.highvolume.exception.ConversionException;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.InputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Implementation of the conversion strategy that uses Apache POI's low-level event-based
 * XML parsing approach for handling very large Excel files with reduced memory usage.
 * <p>
 * This strategy relies on POI's XSSF SAX parsing to stream through the XML content of the Excel file.
 * <p>
 * Contract reference: HighVolumeExcelConverter-Contract-v2.0.1.md (section 7.2)
 *
 * @pre Configuration must be valid and the input file must exist and be readable
 * @post Excel file is processed with minimal memory usage and data is written to output
 * @invariant Memory consumption is kept lower than full-load strategies
 */
public class UserModeEventConversionStrategy implements ConversionStrategy {
    private static final Logger LOG = LoggerFactory.getLogger(UserModeEventConversionStrategy.class);
    
    // Memory and batch processing configuration
    private static final int DEFAULT_CHECKPOINT_ROWS = 50000; // Process in batches of 50,000 rows
    private static final double MEMORY_THRESHOLD_PERCENTAGE = 0.7; // 70% memory usage triggers checkpoint
    private static final long FORCE_GC_INTERVAL_MS = 30000; // Force GC every 30 seconds
    
    /**
     * Executes the conversion process using POI's event-based SAX parsing for XSSF.
     *
     * @param config The converter configuration
     * @pre Input file must exist and be a valid .xlsx file
     * @post Data is written to the output according to the configuration
     * @throws Exception if any conversion errors occur
     */
    @Override
    public void convert(ConverterConfig config) throws Exception {
        LOG.info("Starting conversion using UserModeEventConversionStrategy for file: {}", config.inputFile);
        LOG.info("Target format: {}, Sheet name: {}, Sheet index: {}", 
                config.format, config.sheetName, config.sheetIndex);
        
        // Configure POI for large files
        try {
            ZipSecureFile.setMinInflateRatio(0.01); // Adjusted from 0.001 to a more moderate 0.01
            ZipSecureFile.setMaxEntrySize(6_442_450_944L); // Set max entry size to 6GB to handle large sheet XML
            LOG.info("POI security settings applied: minInflateRatio=0.01, maxEntrySize=6GB");
        } catch (Exception e) {
            LOG.warn("Could not set ZipSecureFile security parameters: {}", e.getMessage());
        }

        OPCPackage pkg = null;
        try {
            try {
                pkg = OPCPackage.open(config.inputFile.toFile(), PackageAccess.READ);
                LOG.info("OPCPackage opened successfully for: {}", config.inputFile);
            } catch (Exception e) {
                LOG.error("Failed to open OPCPackage for file: {}", config.inputFile, e);
                throw new ConversionException("Failed to open Excel file", e);
            }

            processWithUserModeEventApi(pkg, config);
            
            LOG.info("Conversion completed successfully");
        } finally {
            // Ensure OPCPackage is properly closed
            if (pkg != null) {
                try {
                    pkg.close();
                    LOG.info("OPCPackage closed successfully");
                } catch (Exception e) {
                    LOG.warn("Error closing OPCPackage: {}", e.getMessage());
                }
            }
            
            // Consider removing explicit System.gc() calls, as JVM often handles GC better.
            // System.gc();
            // LOG.info("Garbage collection requested to free up memory");
        }
    }

    private void processWithUserModeEventApi(OPCPackage pkg, 
                                         ConverterConfig config) throws Exception {
        XSSFReader xssfReader = new XSSFReader(pkg);
        StylesTable styles = xssfReader.getStylesTable();
        
        // Use ReadOnlySharedStringsTable for better memory usage
        ReadOnlySharedStringsTable sharedStrings;
        try {
            sharedStrings = new ReadOnlySharedStringsTable(pkg);
            LOG.info("ReadOnlySharedStringsTable loaded successfully");
        } catch (SAXException e) {
            LOG.error("Error loading shared strings table: {}", e.getMessage());
            throw new IOException("Failed to load shared strings", e);
        }
        
        String targetSheetName = null;
        int targetSheetIndex = -1;
        InputStream sheetStreamToProcess = null;

        List<String> availableSheetNames = new ArrayList<>();
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        while (sheetIterator.hasNext()) {
            try (InputStream stream = sheetIterator.next()) { // Ensure streams are closed if not used
                availableSheetNames.add(sheetIterator.getSheetName());
            }
        }

        if (availableSheetNames.isEmpty()) {
            throw new ConversionException("No sheets found in the Excel file.");
        }

        if (config.sheetName != null) {
            boolean found = false;
            for (int i = 0; i < availableSheetNames.size(); i++) {
                if (availableSheetNames.get(i).equalsIgnoreCase(config.sheetName)) {
                    targetSheetIndex = i;
                    targetSheetName = availableSheetNames.get(i);
                    found = true;
                    break;
                }
            }
            if (!found) {
                throw new ConversionException("Sheet with name '" + config.sheetName + "' not found. Available sheets: " + availableSheetNames);
            }
        } else if (config.sheetIndex != null) {
            if (config.sheetIndex < 0 || config.sheetIndex >= availableSheetNames.size()) {
                throw new ConversionException("Invalid sheet index: " + config.sheetIndex +
                        ". File contains " + availableSheetNames.size() + " sheets. Available sheets: " + availableSheetNames);
            }
            targetSheetIndex = config.sheetIndex;
            targetSheetName = availableSheetNames.get(targetSheetIndex);
        } else {
            // Default to the first sheet
            targetSheetIndex = 0;
            targetSheetName = availableSheetNames.get(0);
        }

        LOG.info("Target sheet identified: '{}' (Index: {}). Total sheets available: {}", targetSheetName, targetSheetIndex, availableSheetNames.size());

        // Re-iterate to get the target sheet's InputStream
        // This is necessary because the streams from the first iteration were closed by try-with-resources
        XSSFReader.SheetIterator streamSheetIterator = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int currentIndex = 0;
        while (streamSheetIterator.hasNext()) {
            InputStream currentStream = streamSheetIterator.next();
            if (currentIndex == targetSheetIndex) {
                sheetStreamToProcess = currentStream;
                // Note: sheetName from streamSheetIterator.getSheetName() should match targetSheetName
                break;
            }
            // Close streams of non-target sheets immediately to save resources
            try {
                currentStream.close();
            } catch (IOException e) {
                LOG.warn("Error closing non-target sheet stream at index {}: {}", currentIndex, e.getMessage());
            }
            currentIndex++;
        }

        if (sheetStreamToProcess == null) {
            // This should ideally not happen if logic above is correct
            throw new ConversionException("Failed to obtain input stream for target sheet: " + targetSheetName);
        }

        LOG.info("Processing sheet: '{}' (Index: {}) using its dedicated InputStream.", targetSheetName, targetSheetIndex);

        try (InputStream finalSheetStream = sheetStreamToProcess) { // Ensure target stream is closed after processing
            ExcelRowListener rowListener = new ExcelRowListener(config);
            
            final AtomicInteger rowCounter = new AtomicInteger(0);
            final long startTime = System.currentTimeMillis();
            final Runtime runtime = Runtime.getRuntime(); // For memory monitoring
            final Map<Integer, String> currentRowMap = new TreeMap<>(); // Re-initialize per row
            
            SheetContentsHandler sheetHandler = new SheetContentsHandler() {
                private long lastGcLogTime = System.currentTimeMillis();
                
                @Override
                public void startRow(int rowNum) {
                    currentRowMap.clear();
                    // Memory and GC logic from original, consider making it less aggressive or configurable
                    if (rowNum > 0 && rowNum % 1000 == 0) { // Check every 1000 rows
                        long usedMemory = runtime.totalMemory() - runtime.freeMemory();
                        double memoryUsagePercent = (double) usedMemory / runtime.maxMemory();
                        if (memoryUsagePercent > MEMORY_THRESHOLD_PERCENTAGE) {
                            LOG.warn("Memory usage high ({}%) at row {}. Consider increasing heap or optimizing.",
                                    String.format("%.1f", memoryUsagePercent * 100), rowNum);
                            // Explicit GC can be problematic, JVM is often better at deciding when to GC.
                            // System.gc();
                        }
                        long currentTime = System.currentTimeMillis();
                        if (currentTime - lastGcLogTime > FORCE_GC_INTERVAL_MS) {
                            LOG.debug("Periodic check at row {}. Processing rate: {} rows/sec",
                                    rowNum,
                                    String.format("%.1f", rowNum / ((currentTime - startTime) / 1000.0)));
                            // System.gc();
                            lastGcLogTime = currentTime;
                        }
                    }
                }
                
                @Override
                public void endRow(int rowNum) {
                    try {
                        if (rowNum == config.headerRow) {
                            Map<Integer, String> headerMap = new TreeMap<>(currentRowMap);
                            rowListener.invokeHeadMap(headerMap, null);
                            LOG.debug("Processed header row ({}): {}", rowNum, headerMap);
                        } else if (rowNum > config.headerRow) {
                            Map<Integer, String> dataMap = new TreeMap<>(currentRowMap);
                            rowListener.invoke(dataMap, null);
                            int currentRowCount = rowCounter.incrementAndGet();
                            if (currentRowCount % DEFAULT_CHECKPOINT_ROWS == 0) {
                                LOG.info("Reached checkpoint at data row count {}. Excel row {}. Processing rate: {} rows/sec",
                                        currentRowCount, rowNum,
                                        String.format("%.1f", currentRowCount / ((System.currentTimeMillis() - startTime) / 1000.0)));
                                try {
                                    rowListener.flushWriter();
                                } catch (Exception e) {
                                    LOG.warn("Error flushing writer at checkpoint: {}", e.getMessage());
                                }
                            }
                        }
                    } catch (Exception e) {
                        LOG.error("Error processing Excel row {}: {}", rowNum, e.getMessage(), e);
                        if (!config.continueOnError) {
                           // To propagate the error, it should be rethrown or handled to stop processing.
                           // For now, just logging, but this means processing continues.
                           // Consider a flag or a custom exception to signal stopping.
                           throw new RuntimeException("Unrecoverable error processing row " + rowNum, e); 
                        }
                    }
                }
                
                @Override
                public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                    if (cellReference != null) {
                        CellReference ref = new CellReference(cellReference);
                        currentRowMap.put(Integer.valueOf(ref.getCol()), formattedValue != null ? formattedValue : "");
                    }
                }
                
                @Override
                public void headerFooter(String text, boolean isHeader, String tagName) {
                    // Typically ignore header/footer for data extraction
                }
            };
            
            XMLReader sheetParser;
            try {
                sheetParser = XMLReaderFactory.createXMLReader();
            } catch (SAXException e) {
                LOG.error("Could not create XMLReader: {}", e.getMessage(), e);
                throw new ConversionException("Failed to initialize XML parser", e);
            }
            ContentHandler handler = new XSSFSheetXMLHandler(styles, sharedStrings, sheetHandler, false);
            sheetParser.setContentHandler(handler);
            
            try {
                sheetParser.parse(new InputSource(finalSheetStream));
                LOG.info("Sheet processing completed: '{}', total data rows processed by listener: {}",
                        targetSheetName, rowListener.getTotalRowsSuccessfullyWritten());
            } catch (Exception e) {
                LOG.error("Error processing sheet '{}': {}", targetSheetName, e.getMessage(), e);
                throw new ConversionException("Error processing sheet: " + targetSheetName, e);
            }
            
            // Finalize writer
            if (rowListener != null) {
                try {
                    rowListener.doAfterAllAnalysed(null);
                } catch (Exception e) {
                    LOG.warn("Error finalizing rowListener: {}", e.getMessage());
                }
            }
        } catch (Exception e) {
            LOG.error("Error during sheet processing operation for target sheet '{}': {}", targetSheetName, e.getMessage(), e);
            // Ensure ConversionException is thrown to be caught by the main convert() method's error handling
            if (e instanceof ConversionException) throw e;
            throw new ConversionException("Failed to process target sheet: " + targetSheetName, e);
        }
        // Removed System.gc() call between sheets as we process only one sheet with this strategy.
    }
} 