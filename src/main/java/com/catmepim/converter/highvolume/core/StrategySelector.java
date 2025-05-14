package com.catmepim.converter.highvolume.core;

import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.catmepim.converter.highvolume.strategy.ConversionStrategy;
import com.catmepim.converter.highvolume.strategy.StreamingConversionStrategy;
import com.catmepim.converter.highvolume.strategy.UserModeEventConversionStrategy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;

/**
 * Selects the appropriate {@link ConversionStrategy} based on the configuration and file properties.
 * 
 * This class and its methods adhere to the contract defined in:
 * java_tools/high-volume-excel-converter/HighVolumeExcelConverter-Contract-v2.0.1.md
 */
public class StrategySelector {

    private static final Logger logger = LoggerFactory.getLogger(StrategySelector.class);
    // Threshold from contract: 500MB. For files larger, UserModeEvent/Streaming is preferred.
    private static final long LARGE_FILE_THRESHOLD_BYTES = 500L * 1024L * 1024L; // Explicitly long

    /**
     * Selects a conversion strategy based on the provided configuration.
     *
     * @param config The converter configuration.
     * @return The selected {@link ConversionStrategy}.
     * @pre {@code config != null} and has been validated.
     * @post A non-null {@link ConversionStrategy} instance is returned.
     * @invariant The selection logic aims to balance performance and memory usage according to the contract's guidelines.
     */
    public ConversionStrategy selectStrategy(ConverterConfig config) {
        ConversionStrategy selectedStrategy;

        // Priority 1: User-specified strategy hint
        if (config.strategyHint != null && config.strategyHint != ConverterConfig.Strategy.AUTO) {
            switch (config.strategyHint) {
                case STREAMING:
                    logger.info("User hint provided: Selecting StreamingConversionStrategy.");
                    selectedStrategy = new StreamingConversionStrategy();
                    break;
                case USER_MODEL_EVENT:
                    logger.info("User hint provided: Selecting UserModeEventConversionStrategy.");
                    selectedStrategy = new UserModeEventConversionStrategy();
                    break;
                // AUTO case is handled by falling through to selectDefaultStrategy
                default: // Should not happen if AUTO is handled, but as a safeguard
                    logger.warn("Unknown or AUTO strategy hint '{}' encountered directly in switch. Falling back to default selection logic.", config.strategyHint);
                    selectedStrategy = selectDefaultStrategy(config);
                    break;
            }
        } else {
            // Priority 2: Default selection logic (e.g., based on file size or if hint is AUTO)
            if (config.strategyHint == ConverterConfig.Strategy.AUTO) {
                logger.info("Strategy hint is AUTO. Proceeding with default selection logic.");
            } else {
                logger.info("No specific strategy hint provided. Proceeding with default selection logic.");
            }
            selectedStrategy = selectDefaultStrategy(config);
        }
        
        return selectedStrategy;
    }

    /**
     * Selects a default conversion strategy based on file size or other internal heuristics.
     * This method is called if no specific strategy hint is provided or if the hint is AUTO.
     *
     * @param config The converter configuration, assumed to be valid.
     * @return A non-null {@link ConversionStrategy} instance.
     * @pre {@code config != null} and its {@code inputFile} path is set.
     *      While {@code inputFile} should ideally exist and be readable for an accurate size check,
     *      the method has a fallback if the file is not found at selection time.
     * @post A {@link ConversionStrategy} (either {@link UserModeEventConversionStrategy} or 
     *       {@link StreamingConversionStrategy}) is returned.
     */
    private ConversionStrategy selectDefaultStrategy(ConverterConfig config) {
        File inputFile = config.inputFile.toFile(); // Use .toFile() for Path to File conversion
        if (!inputFile.exists()) {
            logger.error("Input file does not exist: {}. Cannot determine size for strategy selection.", config.inputFile);
            // Fallback or throw? For now, let's default to a general strategy and let later stages handle file not found.
            // This part of the logic assumes the file *should* exist for size check.
            // Contract implies pre-validation of file existence, but a check here is safer for strategy selection.
            logger.warn("Input file not found at path: {}. Defaulting to StreamingConversionStrategy.", config.inputFile);
            return new StreamingConversionStrategy(); 
        }
        long fileSize = inputFile.length(); // Gets file size in bytes

        // Contract: "For files expected to be very large (e.g., >500MB), 
        // the User Mode Event API (with XSSFReader and custom SheetContentsHandler) is generally preferred."
        // The StreamingConversionStrategy here is intended to be the EasyExcel based one or similar, which is good for very large files.
        // UserModeEventConversionStrategy is POI's event model.
        // Let's assume StreamingConversionStrategy (EasyExcel or advanced POI streaming) is for the largest files / memory critical.
        // And UserModeEventConversionStrategy (POI event model) is the robust default for large files.

        if (fileSize > LARGE_FILE_THRESHOLD_BYTES) {
            logger.info("File size ({} MB) exceeds threshold ({} MB). Selecting UserModeEventConversionStrategy.", 
                        fileSize / (1024L * 1024L), LARGE_FILE_THRESHOLD_BYTES / (1024L * 1024L)); // Explicitly long for division
            return new UserModeEventConversionStrategy();
        } else {
            // For smaller files, or if the primary streaming strategy is more robust or generally preferred by design.
            // The contract states: "The primary strategy for handling large files will be streaming...". 
            // This could imply StreamingConversionStrategy should be a default if not for the size hint for UserModeEvent.
            // Let's default to StreamingConversionStrategy for files not explicitly large, assuming it's well-optimized.
            logger.info("File size ({} MB) is within threshold or file was not found. Selecting StreamingConversionStrategy as default.", 
                        fileSize / (1024L * 1024L)); // Explicitly long for division
            return new StreamingConversionStrategy();
        }
        // TODO: Further refine strategy selection. For example, if UserModeEvent is considered more robust 
        // or has better feature coverage for certain .xlsx files, it might be preferred even for smaller files.
        // The current logic: >500MB -> UserModeEvent, <=500MB -> Streaming. This aligns with contract's explicit mention for >500MB.
    }
} 