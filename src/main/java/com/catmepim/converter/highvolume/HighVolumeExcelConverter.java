package com.catmepim.converter.highvolume;

import java.util.concurrent.TimeUnit;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.catmepim.converter.highvolume.core.StrategySelector;
import com.catmepim.converter.highvolume.strategy.ConversionStrategy;

import picocli.CommandLine;

/**
 * Main class for the High Volume Excel Converter application.
 * <p>
 * Responsible for parsing command line arguments, setting up the configuration,
 * selecting the appropriate conversion strategy, and executing it.
 * <p>
 * <b>Contract reference:</b> HighVolumeExcelConverter-Contract-v2.0.1.md (см. раздел 14)
 * <p>
 * <b>Recommended JVM Settings:</b>
 * <pre>
 * -XX:+UseG1GC -XX:MaxGCPauseMillis=200 -XX:ConcGCThreads=2 -Xms1g -Xmx4g
 * </pre>
 *
 * @invariant The application attempts to process an Excel file according to user-provided
 *      parameters and the selected strategy. Global settings like ZipSecureFile.minInflateRatio
 *      are applied before processing.
 *      The configuration (ConverterConfig) is immutable after parsing.
 *      All resources (streams, writers) are closed in finally.
 */
public class HighVolumeExcelConverter {

    private static final Logger logger = LoggerFactory.getLogger(HighVolumeExcelConverter.class);

    /**
     * Main entry point for the application.
     *
     * @param args Command line arguments.
     * @pre args != null; args contains valid command-line options as defined in {@link ConverterConfig}.
     * @post The application processes the Excel file or exits with an error code/message.
     *       If successful, output files are generated as specified.
     *       Metrics about the conversion (time, rows, etc.) are logged.
     * @invariant Configuration is not mutated after parsing. All resources are closed in finally.
     * @throws IllegalArgumentException if args is null.
     * @throws Exception if any error occurs during conversion (see contract for details).
     * @contract HighVolumeExcelConverter-Contract-v2.0.1.md
     */
    public static void main(String[] args) {
        if (args == null) {
            throw new IllegalArgumentException("args must not be null");
        }
        long startTime = System.nanoTime();
        logger.info("High Volume Excel Converter starting...");

        ConverterConfig config = new ConverterConfig();
        CommandLine cmd = new CommandLine(config);

        try {
            cmd.parseArgs(args);

            if (cmd.isUsageHelpRequested()) {
                cmd.usage(System.out);
                return;
            }
            if (cmd.isVersionHelpRequested()) {
                cmd.printVersionHelp(System.out);
                return;
            }

            config.validate();
            logger.info("Configuration loaded and validated successfully.");
            logger.info("Input file: {}", config.inputFile);
            logger.info("Output format: {}", config.format);
            if (config.outputPath != null) {
                logger.info("Output path: {}", config.outputPath);
            }
            logger.info("Strategy hint: {}", config.strategyHint);

            ZipSecureFile.setMinInflateRatio(config.minInflateRatio);
            logger.info("ZipSecureFile.minInflateRatio set to: {}", config.minInflateRatio);

            StrategySelector strategySelector = new StrategySelector();
            ConversionStrategy strategy = strategySelector.selectStrategy(config);
            logger.info("Selected strategy: {}", strategy.getClass().getSimpleName());

            logger.info("Starting conversion process...");
            strategy.convert(config);
            logger.info("Conversion process completed successfully.");

        } catch (CommandLine.MissingParameterException ex) {
            logger.error("Missing required parameter(s): {}", ex.getMessage());
            cmd.usage(System.err);
            System.exit(cmd.getCommandSpec().exitCodeOnInvalidInput());
        } catch (CommandLine.ParameterException ex) {
            logger.error("Invalid parameter(s): {}", ex.getMessage());
            cmd.usage(System.err);
            System.exit(cmd.getCommandSpec().exitCodeOnInvalidInput());
        } catch (IllegalArgumentException ex) {
            logger.error("Configuration validation failed: {}", ex.getMessage());
            System.exit(1);
        } catch (Exception ex) {
            logger.error("An unexpected error occurred during conversion: {}", ex.getMessage(), ex);
            System.exit(1);
        } finally {
            long endTime = System.nanoTime();
            long durationMillis = TimeUnit.NANOSECONDS.toMillis(endTime - startTime);
            logger.info("High Volume Excel Converter finished in {} ms.", durationMillis);
            // TODO: Add more detailed metrics (rows processed, output file size, etc.)
            // @post: All resources (streams, writers) must be closed here (see contract §4, §5)
        }
    }
} 