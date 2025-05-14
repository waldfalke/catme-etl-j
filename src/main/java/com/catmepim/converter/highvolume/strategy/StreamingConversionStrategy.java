package com.catmepim.converter.highvolume.strategy;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.catmepim.converter.highvolume.core.ExcelRowListener;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.InputStream;

/**
 * Implements the {@link ConversionStrategy} using a streaming approach, primarily leveraging EasyExcel.
 * This strategy is typically suited for very large Excel files as it aims to minimize memory footprint
 * by processing the file in a stream.
 *
 * <b>Contract reference:</b> HighVolumeExcelConverter-Contract-v2.0.1.md (см. раздел 14)
 *
 * @invariant Memory usage is kept relatively low by processing data in streams via EasyExcel.
 *            The order of rows from the input Excel is preserved in the output.
 *            Configuration (ConverterConfig) is not mutated during execution.
 * @contract HighVolumeExcelConverter-Contract-v2.0.1.md
 */
public class StreamingConversionStrategy implements ConversionStrategy {

    private static final Logger logger = LoggerFactory.getLogger(StreamingConversionStrategy.class);

    /**
     * @pre {@code config != null} and all its relevant fields for this strategy are populated and validated.
     *      Specifically, {@code config.inputFile} must point to a valid, readable Excel file.
     *      {@code config.format} must be specified.
     *      {@code config.tempDir} must be a writable path if used by EasyExcel for temporary files.
     *      Preconditions are checked explicitly and throw IllegalArgumentException if violated.
     * @post The input Excel file is converted to the specified output format using EasyExcel's streaming capabilities 
     *       via the {@link ExcelRowListener}.
     *       If successful, output is generated according to {@code config.outputPath} (for JSON/NDJSON) or
     *       in chunks in {@code config.tempDir} (for CSV).
     *       If an error occurs, an exception is thrown and logged.
     *       The total number of successfully written rows is logged.
     * @invariant Memory usage is kept relatively low by processing data in streams via EasyExcel.
     *            The order of rows from the input Excel is preserved in the output.
     *            Configuration (ConverterConfig) is not mutated during execution.
     * @throws IllegalArgumentException if config or required fields are null/invalid.
     * @throws Exception if any error occurs during conversion (see contract for details).
     * @contract HighVolumeExcelConverter-Contract-v2.0.1.md
     */
    @Override
    public void convert(ConverterConfig config) throws Exception {
        if (config == null) {
            throw new IllegalArgumentException("config must not be null");
        }
        if (config.inputFile == null) {
            throw new IllegalArgumentException("inputFile must not be null");
        }
        if (config.format == null) {
            throw new IllegalArgumentException("format must not be null");
        }
        if (config.tempDir == null) {
            throw new IllegalArgumentException("tempDir must not be null");
        }
        logger.info("Executing Streaming Conversion Strategy for file: {}", config.inputFile);
        logger.info("Output format: {}", config.format);
        logger.info("Header row index: {}", config.headerRow);
        // Note: ZipSecureFile.setMinInflateRatio is set globally in HighVolumeExcelConverter.main

        InputStream inputStream = null;
        ExcelReader excelReader = null;
        ExcelRowListener rowListener = null;

        try {
            inputStream = new FileInputStream(config.inputFile.toFile());

            rowListener = new ExcelRowListener(config);

            excelReader = EasyExcel.read(inputStream, null, rowListener)
                                   .headRowNumber(config.headerRow + 1)
                                   .autoCloseStream(false)
                                   .build();

            ReadSheet readSheet;
            if (config.sheetIndex != null) {
                logger.info("Reading sheet by index: {}", config.sheetIndex);
                readSheet = EasyExcel.readSheet(config.sheetIndex).build();
            } else if (config.sheetName != null && !config.sheetName.trim().isEmpty()) {
                logger.info("Reading sheet by name: {}", config.sheetName);
                readSheet = EasyExcel.readSheet(config.sheetName).build();
            } else {
                logger.info("No sheet specified, reading first sheet (index 0).");
                readSheet = EasyExcel.readSheet(0).build();
            }

            logger.info("Starting to read sheet data using ExcelRowListener...");
            excelReader.read(readSheet);
            logger.info("Finished processing sheet with EasyExcel and ExcelRowListener.");
            logger.info("Total rows successfully written by listener: {}", rowListener.getTotalRowsSuccessfullyWritten());

        } finally {
            if (excelReader != null) {
                try {
                    excelReader.finish();
                } catch (Exception e) {
                    logger.warn("Error during ExcelReader finish: {}", e.getMessage());
                }
            }
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (Exception e) {
                    logger.warn("Error closing input stream: {}", e.getMessage());
                }
            }
            logger.info("Streaming Conversion Strategy execution attempt finished for file: {}", config.inputFile);
        }
    }
} 