# Contract for CatmePoiSheetContentsHandler - v1.0.0

**Version:** 1.0.0
**Date:** YYYY-MM-DD
**Status:** Proposed

## 1. Overview

`CatmePoiSheetContentsHandler` implements Apache POI's `XSSFSheetXMLHandler.SheetContentsHandler` interface. It is a core component of the `UserModeEventConversionStrategy`, responsible for processing SAX events generated from an Excel sheet's XML.

Its primary role is to listen for cell data, assemble rows, distinguish header rows from data rows, format cell content into strings, and then pass this structured row data to an `ExcelRowListener` for further processing and writing.

## 2. Interface

`CatmePoiSheetContentsHandler` implements `org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler`.

```java
package com.catmepim.converter.highvolume.core.poi;

import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.catmepim.converter.highvolume.core.ExcelRowListener;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.util.Map;
import java.util.TreeMap;

public class CatmePoiSheetContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    public CatmePoiSheetContentsHandler(ExcelRowListener rowListener, ConverterConfig config, 
                                        StylesTable stylesTable, ReadOnlySharedStringsTable sharedStringsTable);

    @Override
    public void startRow(int rowNum);

    @Override
    public void endRow(int rowNum);

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment);

    // Optional: Implement if needed, or leave as default (no-op from a base class if available)
    // public void headerFooter(String text, boolean isHeader, String tagName);

    @Override
    public void endSheet();
}
```

## 3. Responsibilities

*   **Row Assembly:** Collect individual cell data (`formattedValue` from the `cell` method) and assemble them into a representation of a row (e.g., `Map<Integer, String>` where Integer is the 0-based column index).
*   **Header Detection:** Identify the header row based on `config.headerRow` (0-indexed).
*   **Data Formatting:** Utilize `org.apache.poi.ss.usermodel.DataFormatter` and `StylesTable` to correctly format cell values (especially dates and numbers) into strings. The `formattedValue` provided to the `cell` method by `XSSFSheetXMLHandler` (if configured with a `DataFormatter`) often handles this.
*   **Shared Strings Resolution:** The `XSSFSheetXMLHandler` typically handles shared string resolution before calling `cell()` if provided with a `SharedStringsTable`.
*   **Column Indexing:** Derive 0-based column indices from cell references (e.g., "A1", "B1" -> 0, 1).
*   **Delegation to `ExcelRowListener`:**
    *   When a header row is completed (at `endRow`), pass the assembled header map to `excelRowListener.invokeHeadMap(headerMap, null)`.
    *   When a data row is completed, pass the assembled data map to `excelRowListener.invoke(dataMap, null)`.
    *   (Note: `AnalysisContext` for `ExcelRowListener` will be null, as it's an EasyExcel specific concept. `ExcelRowListener` must be able_to handle this, or this handler might need to interact with an `IDataWriter` more directly if `ExcelRowListener` proves too tightly coupled.)
*   **Sheet End Signaling:** When `endSheet()` is called, signal to the `ExcelRowListener` that processing of the current sheet is complete (e.g., by calling `excelRowListener.doAfterAllAnalysed(null)` or a dedicated finalization method on the listener), so it can close its `IDataWriter`.

## 4. Dependencies

*   `com.catmepim.converter.highvolume.core.ExcelRowListener`: The listener to which processed rows are delegated.
*   `com.catmepim.converter.highvolume.config.ConverterConfig`: To access configuration like `headerRow`.
*   `org.apache.poi.xssf.model.StylesTable`: For cell style information, used by `DataFormatter`.
*   `org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable`: For resolving shared strings (though `XSSFSheetXMLHandler` might do the lookup).
*   `org.apache.poi.ss.usermodel.DataFormatter`: For converting cell values to strings based on their format.
*   `org.apache.poi.ss.util.CellReference`: Utility to parse cell references like "A1" into row and column indices.
*   `org.slf4j.Logger`: For logging.

## 5. Detailed Method Behavior

### 5.1. `public CatmePoiSheetContentsHandler(ExcelRowListener rowListener, ConverterConfig config, StylesTable stylesTable, ReadOnlySharedStringsTable sharedStringsTable)`

*   **@pre** All arguments are non-null.
*   **@post** The handler is initialized with references to the listener, config, styles, and shared strings table.
*   **@post** Internal state for the current row (e.g., `currentRowMap`, `currentRowNum`) is initialized.
*   **@post** A `DataFormatter` instance is created.

### 5.2. `public void startRow(int rowNum)`

*   **@pre** `rowNum >= 0`.
*   **@post** `this.currentRowNum` is set to `rowNum`.
*   **@post** `this.currentRowMap` (e.g., a `TreeMap<Integer, String>`) is cleared and prepared for new cell data.
*   **@post** Logs the start of a new row, especially if verbose logging is enabled.

### 5.3. `public void endRow(int rowNum)`

*   **@pre** `rowNum == this.currentRowNum`.
*   **@pre** Cells for the current row have been processed by `cell()`.
*   **@post** If `rowNum == config.headerRow`:
    *   `excelRowListener.invokeHeadMap(this.currentRowMap, null)` is called.
    *   Logs header processing.
*   **@post** If `rowNum > config.headerRow` (or other conditions for data rows):
    *   `excelRowListener.invoke(this.currentRowMap, null)` is called.
    *   Logs data row processing.
*   **@post** `this.currentRowMap` may be cleared or reset after processing.
*   **@throws** `RuntimeException` or propagates exceptions from `excelRowListener` if errors occur during delegation and are not handled internally.

### 5.4. `public void cell(String cellReference, String formattedValue, XSSFComment comment)`

*   **@pre** `cellReference != null`. `formattedValue` can be null for empty cells.
*   **@pre** Handler is currently processing a row (i.e., `startRow` has been called for the current row).
*   **@post** The `formattedValue` is stored in `this.currentRowMap` against its 0-based column index.
*   **@post** Column index is derived from `cellReference` (e.g., using `new CellReference(cellReference).getCol()`).
*   **@post** Null `formattedValue` might be stored as null or an empty string, based on desired behavior for empty cells.
*   **@post** `comment` is ignored for now.

### 5.5. `public void endSheet()`

*   **@pre** All rows and cells for the current sheet have been processed.
*   **@post** `excelRowListener.doAfterAllAnalysed(null)` is called to signal the end of sheet processing to the listener, allowing it to finalize its operations (e.g., close the `IDataWriter`).
*   **@post** Logs the end of sheet processing.
*   **@throws** Propagates exceptions from `excelRowListener` if errors occur during finalization.

## 6. Invariants

*   `this.rowListener`, `this.config`, `this.stylesTable`, `this.sharedStringsTable` are not modified after construction.
*   `this.currentRowNum` always reflects the row number passed to the most recent `startRow` call.
*   `this.currentRowMap` contains data only for the current row being processed between `startRow` and `endRow`.

## 7. Error Handling

*   Exceptions from `CellReference` parsing (e.g., invalid `cellReference`) should be handled (e.g., logged, potentially skipped if `config.continueOnError` is true).
*   Exceptions from `excelRowListener` calls are propagated upwards to the `UserModeEventConversionStrategy` to be handled according to its error policy.
*   The handler itself should be robust to empty or malformed cell data, relying on `DataFormatter` and `XSSFSheetXMLHandler` for initial parsing.

---
This contract defines the behavior of the custom SAX event handler for sheet content. 