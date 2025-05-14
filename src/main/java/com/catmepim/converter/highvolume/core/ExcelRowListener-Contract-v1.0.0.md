## Contract for ExcelRowListener (v1.0.0)

**1. Introduction**
   - Component: `ExcelRowListener`
   - Version: 1.0.0
   - Path: `com.catmepim.converter.highvolume.core.ExcelRowListener`
   - Description: An EasyExcel `AnalysisEventListener` responsible for processing rows from an Excel sheet. It handles header identification, data reception, and delegates writing to format-specific writers based on the provided `ConverterConfig`.

**2. Responsibilities**
   - Initialize appropriate data writer(s) (for CSV, NDJSON, JSON) based on `ConverterConfig`.
   - Receive and store header data from the Excel sheet (if `config.headerRow` is effective).
   - Receive data rows from the Excel sheet.
   - Pass row data to the active writer for formatting and output.
   - Manage the lifecycle of the data writer(s) (open, write, flush, close).
   - Handle CSV chunking logic (creating new files after `config.batchSize` records).
   - Log progress, errors, and summary information.
   - Count processed rows.

**3. Interfaces**
   - Implements `com.alibaba.excel.event.AnalysisEventListener<Map<Integer, String>>`. (Data rows will be received as `Map<Integer, String>` where Integer is the 0-indexed column).
   - Constructor: `public ExcelRowListener(ConverterConfig config)`

**4. Preconditions**
   - **Constructor (`ExcelRowListener(ConverterConfig config)`):**
     - `config` is not null and has been validated.
     - `config.format` is a valid `OutputFormat`.
     - `config.inputFile` path is valid (though file reading itself is done by the strategy).
     - `config.outputPath` (for JSON/NDJSON) or `config.tempDir` (for CSV) must specify writable locations.
     - `config.batchSize > 0`.
     - `config.headerRow >= 0`.
   - **`invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context)` (if headers are processed):**
     - `headMap` contains data from the row identified by `excelReaderBuilder.headRowNumber(...)`.
     - `context` может быть null, если listener вызывается из POI-стратегии (CatmePoiSheetContentsHandler).
   - **`invoke(Map<Integer, String> data, AnalysisContext context)`:**
     - `data` contains cell values for a single data row.
     - `context` может быть null, если listener вызывается из POI-стратегии (CatmePoiSheetContentsHandler).
   - **`doAfterAllAnalysed(AnalysisContext context)`:**
     - `context` может быть null, если listener вызывается из POI-стратегии (CatmePoiSheetContentsHandler).
     - All rows from the targeted sheet have been passed to `invoke` or `invokeHeadMap`.
   - **`onException(Exception exception, AnalysisContext context)`:**
     - `exception` is not null.
     - `context` может быть null, если listener вызывается из POI-стратегии (CatmePoiSheetContentsHandler).

**5. Postconditions**
   - **Constructor:**
     - The listener is initialized with the `ConverterConfig`.
     - The appropriate data writer (e.g., `CsvDataWriter`, `JsonDataWriter`) is instantiated and initialized, ready for writing.
       - For CSV, the first chunk file writer is opened in `config.tempDir`.
       - For NDJSON/JSON, the writer to `config.outputPath` is opened (respecting `config.overwrite`).
   - **`invokeHeadMap`:**
     - Header data is stored internally by the listener.
     - The header row is written to the output if the format requires it (e.g., CSV).
   - **`invoke`:**
     - The row data is written to the current output stream/file via the active data writer.
     - Row count is incremented.
     - For CSV, if `rowCount % config.batchSize == 0`, the current CSV chunk is closed, and a new one is opened.
   - **`doAfterAllAnalysed`:**
     - All open data writers are flushed and closed.
     - A summary (e.g., total rows written) is logged.
   - **`onException`:**
     - The exception is logged.
     - If `config.continueOnError` is false, the exception may be re-thrown (or a flag set) to halt processing.
     - If `config.continueOnError` is true, processing attempts to continue for subsequent rows (if the exception is recoverable at EasyExcel's level for the current row).

**6. Invariants**
   - `config` object remains unchanged during the listener's lifecycle.
   - `rowCount` accurately reflects the number of data rows passed to `invoke`.
   - Only one primary data writer is active at a time, corresponding to `config.format`.
   - Output files are named and located according to the contract and `ConverterConfig`.
   - Row order from the input Excel sheet is preserved in the output.

**7. Error Handling**
   - `IOException`s during writer operations are caught, logged, и могут привести к runtime exception для остановки обработки.
   - Ошибки при парсинге EasyExcel или при вызове из POI-стратегии (CatmePoiSheetContentsHandler) обрабатываются в onException. Listener обязан быть устойчив к context==null.
   - File system errors (e.g., unable to create output files/directories) will lead to exceptions.

**8. Related/Dependencies**
   - CatmePoiSheetContentsHandler-Contract-v1.0.0.md

### 8. Logging & Monitoring (Логирование и мониторинг)

- Все ключевые этапы обработки (старт, завершение, ошибки, прогресс) должны логироваться через SLF4J/Logback.
- Прогресс обработки (каждые N строк, где N=10_000 по умолчанию) должен логироваться с указанием:
    - Количества обработанных строк
    - Времени, затраченного на последние N строк
    - Средней скорости обработки (строк/сек) за этот интервал
- В логах должны быть метки времени, чтобы можно было анализировать производительность и выявлять "тупняки".
- @post: После завершения обработки в логах должны быть итоговые метрики: общее количество строк, общее время, средняя скорость.

#### @post для метода invoke:
- После каждых N строк в логах появляется запись с количеством строк, временем за интервал и скоростью.

#### @post для всей обработки:
- В логах фиксируются итоговые метрики (строки, время, скорость). 