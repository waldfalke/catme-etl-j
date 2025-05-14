package com.catmepim.converter.highvolume.config;

import picocli.CommandLine.Command;
import picocli.CommandLine.Option;
import java.nio.file.Path;

/**
 * Configuration holder and Command definition for the HighVolumeExcelConverter.
 * Uses picocli annotations to define command line arguments.
 *
 * This class structure and its fields correspond to the inputs defined in:
 * java_tools/high-volume-excel-converter/HighVolumeExcelConverter-Contract-v2.0.1.md (Section 6, 14)
 *
 * @invariant inputFile != null && format != null && batchSize > 0
 *           && (format == OutputFormat.CSV || outputPath != null) // outputPath is required unless format is CSV
 *           && configuration is not mutated after validation
 * @contract HighVolumeExcelConverter-Contract-v2.0.1.md
 */
@Command(name = "high-volume-converter",
         mixinStandardHelpOptions = true, // Adds --help and --version options
         version = "High Volume Excel Converter 2.0.1",
         description = "Converts large XLSX files to CSV, NDJSON, or JSON format using streaming.")
public class ConverterConfig implements Runnable /* or Callable<Integer> for exit codes */ {

    // --- Picocli Options --- 
    // Making fields public for direct picocli access, or use setters/getters if preferred

    @Option(names = {"-i", "--input"}, required = true, description = "Path to the input XLSX file.")
    public Path inputFile;
    /**
     * @pre path must not be null and must point to an existing .xlsx file (checked later).
     */

    @Option(names = {"-o", "--output"}, description = "Path for the output file (ndjson/json). Ignored for CSV format (output goes to data/temp/).")
    public Path outputPath;
    /**
     * @pre path must not be null if format is not CSV.
     */

    @Option(names = {"-f", "--format"}, required = true, description = "Target output format: ${COMPLETION-CANDIDATES}")
    public OutputFormat format;
    /**
     * @pre format must be one of 'csv', 'ndjson', 'json'.
     */

    @Option(names = {"-s", "--sheetName"}, description = "Name of the sheet to process. Defaults to the first sheet.")
    public String sheetName = null; // Default to null, logic will handle finding the first sheet
    /**
     * @pre sheetName can be null (use first sheet) or a non-empty string.
     */

    @Option(names = {"-b", "--batchSize"}, description = "Batch size for writing records.")
    public int batchSize = 50000;
    /**
     * @pre batchSize must be > 0 (picocli can validate this too).
     */

    @Option(names = {"-c", "--continueOnError"}, description = "Continue processing on structure/XML parsing errors per row.")
    public boolean continueOnError = false; // Default value
    /**
     * @pre No specific precondition, defaults to false.
     */

    @Option(names = {"--temp-dir"}, description = "Directory for temporary files (e.g., CSV chunks, cached strings).")
    public Path tempDir = Path.of("data", "temp"); // Default value
    /**
     * @pre Path should be writable.
     */

    @Option(names = {"--mem-threshold"}, description = "Memory threshold (MB) before EasyExcel switches to disk-based temp files.")
    public int memoryThresholdMb = 100; // Default value (EasyExcel default)
    /**
     * @pre memoryThresholdMb must be > 0.
     */

    @Option(names = {"--min-inflate-ratio"}, description = "Minimum XML inflation ratio for zip bomb protection (set to 0 to disable). Default: 0.01")
    public double minInflateRatio = 0.01; // POI default
    /**
     * @pre minInflateRatio >= 0.
     */

    @Option(names = {"--sheet-index"}, description = "0-based index of the sheet to process. Overrides --sheetName if both are provided. Defaults to 0.")
    public Integer sheetIndex = null; // Default to null, logic prioritizes this over name
    /**
     * @pre sheetIndex can be null or >= 0.
     */

    @Option(names = {"--header-row"}, description = "0-based index of the row containing headers. Defaults to 0.")
    public int headerRow = 0; // Default value
    /**
     * @pre headerRow >= 0.
     */

    @Option(names = {"--date-format"}, description = "Expected date format string (e.g., yyyy-MM-dd HH:mm:ss). Defaults to common Excel formats.")
    public String dateFormat = null; // Default to null, let POI/EasyExcel handle
    /**
     * @pre dateFormat can be null or a valid format string.
     */

    @Option(names = {"-v", "--verbose"}, description = "Enable verbose logging.")
    public boolean verbose = false; // Default value
    /**
     * @pre No specific precondition, defaults to false.
     */

    @Option(names = {"--overwrite"}, description = "Overwrite output file(s) if they already exist.")
    public boolean overwrite = false; // Default value
    /**
     * @pre No specific precondition, defaults to false.
     */

    @Option(names = {"--strategy-hint"}, description = "Hint for conversion strategy: ${COMPLETION-CANDIDATES}. Defaults to AUTO.")
    public Strategy strategyHint = Strategy.AUTO;
    /**
     * @pre strategyHint can be null (defaults to AUTO) or one of the defined strategies.
     */

    @Option(names = {"--pretty-print"}, description = "Enable pretty printing for JSON output. Increases file size and memory usage.")
    public boolean prettyPrint = false; // Default to false for memory efficiency
    /**
     * @pre No specific precondition, defaults to false.
     */

    // TODO: Add @Option annotations for RowMapper and DataWriter specific configurations if needed
    // @Option(names = "--csv-delimiter", description = "Delimiter for CSV output.")
    // public String csvDelimiter = ","; 

    // --- Constructor --- 
    // Default constructor needed by picocli if not using mixins/factories
    public ConverterConfig() {}

    // --- Core Logic (Runnable/Callable) --- 

    @Override
    public void run() {
        // This method is called by picocli after parsing.
        // We can either put the main conversion logic here or 
        // just use this class as a data holder and call the logic
        // from the main HighVolumeExcelConverter class after parsing.
        // For now, keep logic in HighVolumeExcelConverter.main
        System.out.println("Configuration parsed successfully (example):");
        System.out.println("  Input File: " + inputFile);
        System.out.println("  Output Path: " + outputPath);
        System.out.println("  Format: " + format);
        System.out.println("  Sheet Name: " + (sheetName == null ? "<Default>" : sheetName));
        System.out.println("  Batch Size: " + batchSize);
        System.out.println("  Continue on Error: " + continueOnError);
        System.out.println("  Temp Dir: " + tempDir);
        System.out.println("  Memory Threshold (MB): " + memoryThresholdMb);
        System.out.println("  Min Inflate Ratio: " + minInflateRatio);
        System.out.println("  Sheet Index: " + (sheetIndex == null ? "<Default>" : sheetIndex));
        System.out.println("  Header Row: " + headerRow);
        System.out.println("  Date Format: " + (dateFormat == null ? "<Default>" : dateFormat));
        System.out.println("  Verbose: " + verbose);
        System.out.println("  Overwrite: " + overwrite);
        System.out.println("  Strategy Hint: " + strategyHint);
        System.out.println("  Pretty Print: " + prettyPrint);

        // Basic validation that picocli might not cover easily
        validate(); 
    }

    /**
     * Performs basic validation after picocli populates fields.
     * @throws IllegalArgumentException if validation fails.
     * @pre All fields populated by picocli. All required options are set.
     * @post Configuration is considered valid according to basic rules. Fields are not mutated after validation.
     * @contract HighVolumeExcelConverter-Contract-v2.0.1.md
     */
    public void validate() {
         if (batchSize <= 0) {
            throw new IllegalArgumentException("Batch size must be positive: " + batchSize);
        }
        if (memoryThresholdMb <= 0) {
            throw new IllegalArgumentException("Memory threshold must be positive: " + memoryThresholdMb);
        }
        if (minInflateRatio < 0) {
            throw new IllegalArgumentException("Minimum inflate ratio cannot be negative: " + minInflateRatio);
        }
        if (sheetIndex != null && sheetIndex < 0) {
            throw new IllegalArgumentException("Sheet index cannot be negative: " + sheetIndex);
        }
        if (headerRow < 0) {
            throw new IllegalArgumentException("Header row index cannot be negative: " + headerRow);
        }
        if (format != OutputFormat.CSV && outputPath == null) {
             throw new IllegalArgumentException("Output path (--output) is required for format: " + format);
        }
        // More complex validation (like file existence) can be done later
        // in the main logic before starting the actual processing.
    }


    // Enum for output formats to provide type safety and completion candidates for picocli
    public enum OutputFormat {
        CSV, NDJSON, JSON
    }

    // Enum for conversion strategies
    public enum Strategy {
        STREAMING, // Represents a highly optimized streaming approach (e.g. EasyExcel or similar)
        USER_MODEL_EVENT, // Represents Apache POI's User Model Event API (XSSFReader)
        AUTO // Selector will decide based on file size or other heuristics
    }

} 