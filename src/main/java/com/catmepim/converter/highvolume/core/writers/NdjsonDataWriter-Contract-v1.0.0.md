# Contract for NdjsonDataWriter - v1.0.0

**Version:** 1.0.0
**Date:** YYYY-MM-DD
**Status:** Proposed

## 1. Overview

`NdjsonDataWriter` is an implementation of the `IDataWriter` interface responsible for writing Excel row data into the NDJSON (Newline Delimited JSON) format. Each row from the Excel sheet is converted into a single JSON object, and each JSON object is written to the output file followed by a newline character.

This component is part of the HighVolumeExcelConverter module.

## 2. Interface

`NdjsonDataWriter` implements `com.catmepim.converter.highvolume.core.writers.IDataWriter`.

```java
public class NdjsonDataWriter implements IDataWriter {
    // Constructors
    public NdjsonDataWriter(ConverterConfig config, Map<Integer, String> initialHeaderMap);

    // Interface methods from IDataWriter
    @Override
    public void open() throws IOException;

    @Override
    public void writeHeader(Map<Integer, String> headerData) throws IOException;

    @Override
    public void writeRow(Map<Integer, String> rowData) throws IOException;

    @Override
    public void flush() throws IOException;

    @Override
    public void close() throws IOException;

    @Override
    public long getRowsWrittenCount();
}
```

## 3. Responsibilities

*   Initialize and manage the output file stream for NDJSON.
*   Convert each row of data (represented as `Map<Integer, String>`) into a JSON object string.
*   Write each JSON object string to the output file, ensuring each object is on a new line.
*   Store header information to use as keys for the JSON objects.
*   Handle file system operations like creating the output file and parent directories if necessary.
*   Manage resources (e.g., `BufferedWriter`, `JsonGenerator`) and ensure they are closed properly.
*   Track the count of successfully written rows.

## 4. Dependencies

*   `com.catmepim.converter.highvolume.config.ConverterConfig`: For configuration details like output path.
*   `com.fasterxml.jackson.core.JsonGenerator` (or similar Jackson streaming API): For efficient JSON object generation.
*   `java.io.BufferedWriter`, `java.io.OutputStreamWriter`, `java.nio.file.Files`, `java.nio.file.Path`: For file writing operations.
*   `org.slf4j.Logger`: For logging.

## 5. Detailed Method Behavior (Preconditions, Postconditions, Invariants)

### 5.1. `public NdjsonDataWriter(ConverterConfig config, Map<Integer, String> initialHeaderMap)`

*   **@pre** `config != null`
*   **@pre** `config.outputPath != null` (as NDJSON is not chunked to temp like CSV)
*   **@pre** `initialHeaderMap != null` (can be empty, will be populated by `writeHeader`)
*   **@post** A new `NdjsonDataWriter` instance is created.
*   **@post** `this.config` is initialized with the provided `config`.
*   **@post** `this.orderedHeaders` is initialized (e.g., as an empty `ArrayList`).
*   **@post** `this.rowsWrittenCount` is initialized to 0.
*   **@post** `this.objectMapper` (if used directly) or `this.jsonFactory` is initialized.

### 5.2. `public void open() throws IOException`

*   **@pre** `this.config.outputPath` refers to a valid and writable file path.
*   **@pre** If `config.overwrite` is false, the file at `config.outputPath` must not exist, or an `IOException` (e.g., `FileAlreadyExistsException`) should be thrown.
*   **@pre** Parent directories for `config.outputPath` either exist or can be created.
*   **@post** The output file at `config.outputPath` is opened (or created and opened) for writing.
*   **@post** `this.writer` (e.g., `BufferedWriter`) and `this.jsonGenerator` are initialized and ready for writing.
*   **@post** If `config.overwrite` is true and the file existed, it is truncated or replaced.
*   **@throws** `IOException` if parent directories cannot be created, file cannot be opened/created due to permissions, disk space, or if `overwrite` is false and file exists.

### 5.3. `public void writeHeader(Map<Integer, String> headerData) throws IOException`

*   **@pre** The writer is open (i.e., `open()` has been successfully called).
*   **@pre** `headerData != null`.
*   **@post** `this.orderedHeaders` is populated based on `headerData`, maintaining the order of columns (e.g., by sorting column indices if not already ordered, or using `TreeMap` keys).
*   **@post** No actual data is written to the file by this method for NDJSON, as headers are part of each JSON line. The headers are stored for use in `writeRow`.

### 5.4. `public void writeRow(Map<Integer, String> rowData) throws IOException`

*   **@pre** The writer is open.
*   **@pre** `rowData != null`.
*   **@pre** `this.orderedHeaders` should ideally be populated by a prior `writeHeader` call if headers are to be used as JSON keys. If `orderedHeaders` is empty, column indices might be used as keys (e.g., "0", "1", ...).
*   **@post** A single JSON object representing `rowData` is written to the output file.
*   **@post** The JSON object is followed by a newline character (`\\n`).
*   **@post** `this.rowsWrittenCount` is incremented by 1.
*   **@throws** `IOException` if any error occurs during JSON generation or writing to the file.

### 5.5. `public void flush() throws IOException`

*   **@pre** The writer is open.
*   **@post** All buffered data in `this.jsonGenerator` and `this.writer` is flushed to the underlying file system.
*   **@throws** `IOException` if an error occurs during flushing.

### 5.6. `public void close() throws IOException`

*   **@pre** The writer can be open or already closed (method should be idempotent).
*   **@post** `this.jsonGenerator` (if used) is closed.
*   **@post** `this.writer` (e.g., `BufferedWriter`) is flushed and closed.
*   **@post** All file system resources are released.
*   **@post** The writer is no longer usable for writing.
*   **@throws** `IOException` if an error occurs during closing of resources. Underlying exceptions from `jsonGenerator.close()` or `writer.close()` should be propagated.

### 5.7. `public long getRowsWrittenCount()`

*   **@pre** None.
*   **@post** Returns the total number of rows successfully written by `writeRow`.

## 6. Invariants

*   `this.config` object's relevant fields (e.g., `outputPath`) are not modified after construction.
*   `this.rowsWrittenCount` accurately reflects the number of times `writeRow` completed successfully.
*   If `this.jsonGenerator` or `this.writer` are not null, they refer to a valid, open stream/writer, unless `close()` has been called.

## 7. Error Handling

*   `IOException`s are propagated from file operations or JSON generation.
*   `IllegalArgumentException` might be thrown from the constructor if `config` or `outputPath` is null (though `IDataWriter` users like `ExcelRowListener` should guard this).
*   File existence checks (for `overwrite` flag) are handled in `open()`.

## 8. Data Format Details

*   Each line in the output file will be a complete, self-contained JSON object.
*   Keys for JSON objects will be derived from `orderedHeaders`. If headers are not available or a cell index in a row does not correspond to a known header, a fallback mechanism (e.g., stringified column index) will be used for keys.
*   All cell values are treated as strings by default during JSON generation, unless specific type conversion logic is added (currently out of scope for this writer, which assumes string map input).

## 9. Performance Considerations

*   Uses Jackson's streaming API (`JsonGenerator`) for efficient, low-memory JSON writing.
*   Uses `BufferedWriter` for efficient file I/O.
*   Avoids loading the entire dataset into memory.
---
This contract provides a clear set of expectations for the `NdjsonDataWriter` component. 