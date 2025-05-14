# Contract for UserModeEventConversionStrategy - v1.0.0

**Version:** 1.0.0
**Date:** YYYY-MM-DD
**Status:** Proposed

## 1. Overview

The `UserModeEventConversionStrategy` implements the `ConversionStrategy` interface. It is designed to process very large XLSX files by leveraging Apache POI's SAX-based event model (specifically `XSSFReader` and a custom `XSSFSheetXMLHandler.SheetContentsHandler`). This approach avoids loading the entire workbook into memory, making it suitable for files that exceed the memory capacity of user-model-based approaches (like Apache POI's UserModel or even EasyExcel's default mode for extremely large files).

This strategy focuses on streaming rows from the specified Excel sheet and passing them to an `ExcelRowListener` (or a similar data handling mechanism) which, in turn, uses an appropriate `IDataWriter` to output data in the desired format (CSV, JSON, NDJSON).

## 2. Interface

`UserModeEventConversionStrategy` implements `com.catmepim.converter.highvolume.strategy.ConversionStrategy`.

```java
package com.catmepim.converter.highvolume.strategy;

import com.catmepim.converter.highvolume.core.ConverterConfig;

public class UserModeEventConversionStrategy implements ConversionStrategy {
    public UserModeEventConversionStrategy();

    @Override
    public void convert(ConverterConfig config) throws Exception;
}
```

## 3. Responsibilities

*   Initialize and configure Apache POI's `XSSFReader` for processing the input XLSX file.
*   Identify the target sheet to process based on `ConverterConfig` (by name or index).
*   Set up a custom `XSSFSheetXMLHandler.SheetContentsHandler` to intercept row and cell data from the SAX events.
*   Manage the lifecycle of shared strings table (potentially using `ReadOnlySharedStringsTable` or `SharedStringsTable` with disk caching if POI supports it effectively in this context).
*   Instantiate and utilize an `ExcelRowListener` (or a compatible row processing component) to handle header identification, row data aggregation, and delegation to the appropriate `IDataWriter`.
*   Ensure all resources (file streams, etc.) are properly closed after processing or in case of an error.
*   Handle potential XML parsing errors or structural issues within the Excel file gracefully, respecting the `continueOnError` flag in `ConverterConfig`.

## 4. Dependencies

*   `com.catmepim.converter.highvolume.core.ConverterConfig`: For input file path, sheet selection, output format, and other conversion parameters.
*   `com.catmepim.converter.highvolume.core.ExcelRowListener`: To process discovered headers and data rows, and to manage the `IDataWriter`.
*   Apache POI libraries:
    *   `poi-ooxml` (for `XSSFReader`, `XSSFSheetXMLHandler`, etc.)
    *   `poi-ooxml-lite` (or `poi-ooxml-full` if specific features are needed for event model)
*   `org.xml.sax.XMLReader`: For parsing sheet XML.
*   `org.slf4j.Logger`: For logging.

## 5. Detailed Method Behavior

### 5.1. `public UserModeEventConversionStrategy()`

*   **@pre** None.
*   **@post** A new `UserModeEventConversionStrategy` instance is created.

### 5.2. `public void convert(ConverterConfig config) throws Exception`

*   **@pre** `config != null` and has been validated (e.g., `inputFile` exists, `format` is valid).
*   **@pre** `config.minInflateRatio` is set appropriately for `ZipSecureFile.setMinInflateRatio()` before opening the OPCPackage.
*   **@post** The specified sheet in the input XLSX file is processed.
*   **@post** Data from the sheet is written to the output destination in the format specified by `config.format` via the `ExcelRowListener` and its `IDataWriter`.
*   **@post** All resources are released.
*   **@throws** `Exception` if any critical error occurs during file processing, XML parsing, or data writing that cannot be handled by `continueOnError` logic.
*   **Behavioral Outline:**
    1.  Log strategy initiation.
    2.  Apply global settings like `ZipSecureFile.setMinInflateRatio(config.minInflateRatio)`.
    3.  Open the `OPCPackage` from `config.inputFile`.
    4.  Create an `XSSFReader` from the `OPCPackage`.
    5.  Obtain the `SharedStringsTable` (consider `ReadOnlySharedStringsTable` for memory efficiency).
    6.  Instantiate the `ExcelRowListener` using the `config`. This listener will internally set up the correct `IDataWriter`.
    7.  Create an instance of the custom `PoiSheetContentsHandler` (e.g., `CatmePoiSheetHandler`), passing it the `ExcelRowListener` and `config` (for header row index, etc.).
    8.  Obtain an `XMLReader` (SAX parser).
    9.  Configure the `XSSFSheetXMLHandler` with the `SharedStringsTable`, custom `SheetContentsHandler`, styles table (if needed, though typically not for raw data extraction), and other necessary parameters.
    10. Determine the target sheet's relationship ID (`rId`) based on `config.sheetName` or `config.sheetIndex`. This involves iterating through `XSSFReader.getSheetIterator()`.
    11. If the target sheet is found, get its `InputStream` using `xssfReader.getSheet(rId)`.
    12. Create an `InputSource` from the sheet's `InputStream`.
    13. Start parsing: `xmlReader.parse(inputSource)`.
    14. Ensure `ExcelRowListener` (and its `IDataWriter`) are properly closed in a `finally` block or after processing completes (e.g., by calling methods on `ExcelRowListener` that trigger `dataWriter.close()`).
    15. Close the `OPCPackage`.

## 6. Custom SheetContentsHandler (e.g., `CatmePoiSheetHandler`)

This strategy will require a custom implementation of `XSSFSheetXMLHandler.SheetContentsHandler`. This handler will be responsible for:
*   Identifying the header row based on `config.headerRow`.
*   Buffering cells for a single row.
*   When a row is complete (at `endRow` SAX event):
    *   If it's the header row, pass the collected header data (`Map<Integer, String>`) to `excelRowListener.invokeHeadMap(headers, null)`. (Note: `AnalysisContext` from EasyExcel is not available here; a null or dummy context might be passed if the listener requires it, or the listener adapted).
    *   If it's a data row, pass the collected row data (`Map<Integer, String>`) to `excelRowListener.invoke(rowData, null)`.
*   Handling cell types and formatting them to strings (e.g., numeric, date, boolean cells).
*   Managing cell references (`cellRef`) to determine column indices.

A separate contract for this custom handler should be created.

## 7. Error Handling

*   `IOExceptions` from file operations.
*   `SAXExceptions` from XML parsing.
*   `OpenXML4JExceptions` from `OPCPackage` handling.
*   Exceptions from the `ExcelRowListener` or `IDataWriter` during data processing/writing.
*   The `continueOnError` flag in `ConverterConfig` should be consulted. If true, errors in individual row processing (e.g., bad cell format) might be logged and skipped. Critical errors (e.g., file not found, unreadable sheet) should halt processing.

## 8. Performance and Memory

*   Designed for low memory footprint by streaming XML events.
*   Performance will depend on the SAX parser and the efficiency of the custom `SheetContentsHandler` and the `IDataWriter`.
*   Disk-caching for shared strings (if POI's `SharedStringsTable` is used instead of `ReadOnlySharedStringsTable` and configured for it) might impact performance but save memory.

## 9. Preconditions / Invariants for the class

*   **@inv** The `ConverterConfig` passed to `convert()` is not modified by this strategy.

---
This contract outlines the design for a robust, event-driven Excel processing strategy suitable for very large files. 