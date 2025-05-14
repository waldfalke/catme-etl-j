package com.catmepim.converter.highvolume.strategy;

import com.catmepim.converter.highvolume.config.ConverterConfig;

/**
 * Interface for Excel conversion strategies.
 * Each strategy will implement a specific approach to convert an Excel file.
 * 
 * This interface and its implementations adhere to the contract defined in:
 * java_tools/high-volume-excel-converter/HighVolumeExcelConverter-Contract-v2.0.1.md
 */
public interface ConversionStrategy {

    /**
     * Executes the conversion process based on the provided configuration.
     *
     * @param config The configuration object containing all conversion parameters.
     * @throws Exception if any error occurs during the conversion process.
     * @pre The {@code config} object is not null and has been validated.
     * @post The Excel file specified in {@code config.inputFile} has been processed according to the strategy,
     *       and output files (or data streams) are generated as per {@code config.format} and {@code config.outputPath}.
     *       Or, an exception is thrown detailing the failure.
     */
    void convert(ConverterConfig config) throws Exception;
} 