package com.catmepim.converter.highvolume.core.poi;

import com.catmepim.converter.highvolume.config.ConverterConfig;
import com.catmepim.converter.highvolume.core.ExcelRowListener;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;
import java.util.TreeMap;

/**
 * Custom {@link SheetContentsHandler} for processing XLSX sheet data via SAX events.
 * This handler assembles rows from cell data and delegates them to an {@link ExcelRowListener}.
 *
 * It identifies header rows, formats cell values, and manages the row-building process.
 *
 * Adheres to contract: CatmePoiSheetContentsHandler-Contract-v1.0.0.md
 *
 * @inv {@code this.rowListener}, {@code this.config}, {@code this.stylesTable},
 *      {@code this.sharedStringsTable} are not modified after construction.
 * @inv {@code this.currentRowNum} reflects the row number of the current row being processed.
 * @inv {@code this.currentRowMap} contains data for the current row between startRow and endRow calls.
 */
public class CatmePoiSheetContentsHandler implements SheetContentsHandler {

    private static final Logger LOG = LoggerFactory.getLogger(CatmePoiSheetContentsHandler.class);

    private final ExcelRowListener rowListener;
    private final ConverterConfig config;
    private final StylesTable stylesTable;
    private final ReadOnlySharedStringsTable sharedStringsTable;
    private final DataFormatter dataFormatter;

    private Map<Integer, String> currentRowMap;
    private int currentRowNum = -1;
    private int currentMinColumn = -1;
    private int currentMaxColumn = -1;

    /**
     * Constructs a new CatmePoiSheetContentsHandler.
     *
     * @param rowListener The listener to delegate processed rows to.
     * @param config The converter configuration.
     * @param stylesTable The styles table from the XLSX file.
     * @param sharedStringsTable The shared strings table from the XLSX file.
     * @pre All arguments are non-null.
     * @post Handler is initialized. Internal state for row processing is reset.
     */
    public CatmePoiSheetContentsHandler(ExcelRowListener rowListener, ConverterConfig config,
                                        StylesTable stylesTable, ReadOnlySharedStringsTable sharedStringsTable) {
        this.rowListener = rowListener;
        this.config = config;
        this.stylesTable = stylesTable;
        this.sharedStringsTable = sharedStringsTable;
        this.dataFormatter = new DataFormatter();
        this.currentRowMap = new TreeMap<>();
        LOG.debug("CatmePoiSheetContentsHandler initialized.");
    }

    /**
     * Called at the start of a new row.
     * @param rowNum 0-based row number.
     * @pre {@code rowNum >= 0}.
     * @post Internal state for the current row is prepared.
     */
    @Override
    public void startRow(int rowNum) {
        this.currentRowNum = rowNum;
        this.currentRowMap.clear();
        this.currentMinColumn = -1;
        this.currentMaxColumn = -1;
        LOG.trace("Starting row: {}", rowNum);
    }

    /**
     * Called when a cell is encountered.
     * @param cellReference The cell reference (e.g., "A1", "B2").
     * @param formattedValue The formatted value of the cell as a String, or null for blank cells.
     * @param comment The cell comment, or null if none.
     * @pre Handler is processing a row (startRow called).
     * @post Cell data is added to the current row map.
     */
    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        if (cellReference == null) {
            LOG.warn("Cell reference is null for row {}. Skipping cell.", currentRowNum);
            return;
        }
        CellReference ref = new CellReference(cellReference);
        int colIdx = ref.getCol();

        if (currentMinColumn == -1 || colIdx < currentMinColumn) {
            currentMinColumn = colIdx;
        }
        if (colIdx > currentMaxColumn) {
            currentMaxColumn = colIdx;
        }

        currentRowMap.put(colIdx, formattedValue);
        LOG.trace("  Cell: {}[{}], Value: '{}'", cellReference, colIdx, formattedValue);
    }

    /**
     * Called at the end of a row.
     * @param rowNum 0-based row number.
     * @pre {@code rowNum == this.currentRowNum}.
     * @post The assembled row (header or data) is passed to the ExcelRowListener.
     */
    @Override
    public void endRow(int rowNum) {
        LOG.trace("Ending row: {}", rowNum);
        if (this.currentRowNum != rowNum) {
            LOG.warn("endRow called with rowNum {} but current row is {}. This might indicate a parsing issue or unexpected event order.", rowNum, this.currentRowNum);
        }

        if (currentRowMap.isEmpty() && currentMinColumn == -1) { // Handle truly empty rows that had no cells.
            LOG.trace("Skipping empty row {}.", rowNum);
            return;
        }

        if (currentRowNum == config.headerRow) {
            LOG.debug("Processing header row {}: {}", currentRowNum, currentRowMap);
            try {
                // Pass a defensive copy of the map
                rowListener.invokeHeadMap(new TreeMap<>(currentRowMap), null);
            } catch (Exception e) {
                LOG.error("Error invoking headMap on ExcelRowListener for row {}: {}", currentRowNum, e.getMessage(), e);
                throw new RuntimeException("Error processing header row " + currentRowNum, e);
            }
        } else if (currentRowNum > config.headerRow) {
            LOG.trace("Processing data row {}: {}", currentRowNum, currentRowMap);
            try {
                // Pass a defensive copy of the map
                rowListener.invoke(new TreeMap<>(currentRowMap), null);
            } catch (Exception e) {
                LOG.error("Error invoking invoke on ExcelRowListener for row {}: {}", currentRowNum, e.getMessage(), e);
                if (!config.continueOnError) {
                    throw new RuntimeException("Error processing data row " + currentRowNum, e);
                } else {
                    LOG.warn("Skipping data row {} due to error (continueOnError=true): {}", currentRowNum, e.getMessage());
                }
            }
        } else {
            // Rows before the configured headerRow are ignored.
            LOG.trace("Skipping row {} (it is before configured header row {}).", currentRowNum, config.headerRow);
        }
    }

    /**
     * Called at the end of processing the current sheet.
     * @post Signals to the ExcelRowListener that sheet processing is complete.
     */
    @Override
    public void endSheet() {
        LOG.info("Finished processing sheet. Signaling ExcelRowListener to finalize.");
        try {
            // Passing null for AnalysisContext, as it's specific to EasyExcel and not available here.
            // ExcelRowListener needs to be robust enough to handle a null context for this call,
            // primarily to close its underlying IDataWriter.
            rowListener.doAfterAllAnalysed(null);
        } catch (Exception e) {
            LOG.error("Error calling doAfterAllAnalysed on ExcelRowListener at endSheet: {}", e.getMessage(), e);
            // This is a critical error if the writer cannot be closed properly.
            throw new RuntimeException("Error finalizing ExcelRowListener at end of sheet", e);
        }
    }

    // The headerFooter method is part of the SheetContentsHandler interface but often not needed.
    // Default implementation is no-op.
    // @Override
    // public void headerFooter(String text, boolean isHeader, String tagName) {}
}