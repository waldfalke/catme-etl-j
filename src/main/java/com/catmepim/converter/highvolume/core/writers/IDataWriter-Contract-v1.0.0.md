## Contract for IDataWriter and Implementations (v1.0.0)

**1. Introduction**
   - Component: `IDataWriter` (interface) and its implementations (`CsvDataWriter`, `JsonDataWriter`, `NdjsonDataWriter`)
   - Version: 1.0.0
   - Path:
     - `com.catmepim.converter.highvolume.core.writers.IDataWriter`
     - `com.catmepim.converter.highvolume.core.writers.CsvDataWriter`
     - `com.catmepim.converter.highvolume.core.writers.JsonDataWriter`
     - `com.catmepim.converter.highvolume.core.writers.NdjsonDataWriter`
   - Description: `IDataWriter` defines a common interface for writing processed Excel row data to different output formats. Concrete implementations handle the specifics of each format.

**2. `IDataWriter` Interface**
   - **Methods:**
     - `void open() throws IOException;` (Idempotent, prepares the writer, opens files/streams)
     - `void writeHeader(Map<Integer, String> headerData) throws IOException;` (Writes header if applicable to the format. `headerData` keys are 0-indexed column indices, values are header strings.)
     - `void writeRow(Map<Integer, String> rowData) throws IOException;` (Writes a single data row. `rowData` keys are 0-indexed column indices, values are cell strings.)
     - `void flush() throws IOException;`
     - `void close() throws IOException;` (Flushes and closes all resources.)
     - `long getRowsWrittenCount();` (Returns the number of data rows successfully written by this writer instance.)

**3. `CsvDataWriter` Implementation**
   - **Responsibilities:**
     - Writes data to CSV files.
     - Handles CSV chunking: creates new files in `config.tempDir` named `<original_filename>-chunk-N.csv` after `config.batchSize` rows.
     - Uses Apache Commons CSV (`CSVPrinter`).
   - **Constructor:** `public CsvDataWriter(ConverterConfig config, Map<Integer, String> headerMap)`
     - `headerMap` is provided by `ExcelRowListener` and is used by `CsvDataWriter` if it needs to write headers immediately upon opening the first chunk, or to know the column structure.
   - **Preconditions (Constructor):**
     - `config` is valid, `config.format` is CSV.
     - `config.tempDir` is writable.
     - `config.inputFile` is set (to derive original filename for chunks).
     - `headerMap` is not null (can be empty if no headers).
   - **Postconditions (Methods):**
     - `open()`: Opens the first CSV chunk file for writing.
     - `writeHeader()`: Writes the provided `headerData` as the first line to the current CSV chunk if not already written.
     - `writeRow()`: Writes row data as a CSV record. Increments internal row counter for current chunk. If chunk size is reached, closes current chunk, opens next one, and writes headers to the new chunk.
     - `close()`: Closes the current CSVPrinter and underlying writer.
   - **Invariants:**
     - Output files are in `config.tempDir`.
     - Chunk files are correctly named and sequenced.

**4. `JsonDataWriter` Implementation**
   - **Responsibilities:**
     - Writes data as a single JSON array to `config.outputPath`.
     - Uses Jackson streaming API (`JsonGenerator`).
   - **Constructor:** `public JsonDataWriter(ConverterConfig config, Map<Integer, String> headerMap)`
     - `headerMap` is used to map 0-indexed data row values to JSON object keys.
   - **Preconditions (Constructor):**
     - `config` is valid, `config.format` is JSON.
     - `config.outputPath` is writable.
     - `headerMap` is not null and preferably not empty (for meaningful JSON keys).
   - **Postconditions (Methods):**
     - `open()`: Opens `config.outputPath` for writing, writes the opening array bracket `[`.
     - `writeHeader()`: Stores/updates the internal header map. Does not write to output directly, as JSON headers are per-object.
     - `writeRow()`: Writes row data as a JSON object within the array. Uses stored headers as keys.
     - `close()`: Writes the closing array bracket `]`, closes `JsonGenerator` and underlying writer.
   - **Invariants:**
     - Output is a valid JSON array.
     - Output file is at `config.outputPath`.

**5. `NdjsonDataWriter` Implementation**
   - **Responsibilities:**
     - Writes data as Newline Delimited JSON (NDJSON) to `config.outputPath`. Each row is a separate JSON object on a new line.
     - Uses Jackson streaming API (`JsonGenerator` for each object, or write object by object).
   - **Constructor:** `public NdjsonDataWriter(ConverterConfig config, Map<Integer, String> headerMap)`
     - `headerMap` is used to map 0-indexed data row values to JSON object keys.
   - **Preconditions (Constructor):**
     - `config` is valid, `config.format` is NDJSON.
     - `config.outputPath` is writable.
     - `headerMap` is not null and preferably not empty.
   - **Postconditions (Methods):**
     - `open()`: Opens `config.outputPath` for writing.
     - `writeHeader()`: Stores/updates the internal header map.
     - `writeRow()`: Writes row data as a JSON object, followed by a newline. Uses stored headers as keys.
     - `close()`: Closes underlying writer.
   - **Invariants:**
     - Output is valid NDJSON.
     - Output file is at `config.outputPath`. 