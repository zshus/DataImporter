package dataimporter.implementation.service;

import dataimporter.implementation.model.ExcelCellData;
import dataimporter.implementation.utils.DataImporterRuntimeException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class DataReader implements AutoCloseable {
    private Workbook workbook;
    private Sheet sheet;

    public DataReader(File excelFile) throws IOException {
        if (excelFile == null || !excelFile.exists()) {
            throw new DataImporterRuntimeException("Excel file not found.");
        }
        this.workbook = WorkbookFactory.create(excelFile);
    }

    public boolean hasNextRow(int rowNo) {
        return sheet.getRow(rowNo) != null;
    }

    public List<ExcelCellData> readHeaderRow(int headerRowNo) {
        if (sheet == null) {
            throw new DataImporterRuntimeException("Sheet is null");
        }
        if (!hasNextRow(headerRowNo)) {
            throw new DataImporterRuntimeException("Row number not found");
        }
        try (Stream<Cell> cellStream = StreamSupport.stream(sheet.getRow(headerRowNo).spliterator(), false)) {
            return cellStream.sequential()
                    .map(cell -> {
                        // add column
                        Object rawData = getValue(cell, cell.getCellType());
                        if (rawData != null) {
                            return evaluateCellData(cell, rawData.toString().trim(), null);
                        }
                        return null;
                    })
                    .filter(Objects::nonNull)
                    .collect(Collectors.toList());
        }
    }

    public List<ExcelCellData> readDataRow(int dataRowNo, List<ExcelCellData> headerRowData) {
        var colIndxes = headerRowData.stream().map(ExcelCellData::getColumnIndex).collect(Collectors.toList());
        if (sheet == null) {
            throw new DataImporterRuntimeException("Sheet is null");
        }
        if (!hasNextRow(dataRowNo)) {
            throw new DataImporterRuntimeException("Row number not found");
        }
        try (Stream<Cell> cellStream = StreamSupport.stream(sheet.getRow(dataRowNo).spliterator(), false)) {
            return cellStream.sequential()
                    .map(cell -> {
                        if (DataProcessor.logNode.isTraceEnabled()) {
                            DataProcessor.logNode.trace("Reading excel cell " + getCellName(cell) + " from row " + dataRowNo);
                        }
                        if (!colIndxes.contains(cell.getColumnIndex())) {
                            return null;
                        }
                        // add column
                        Object rawData = getValue(cell, cell.getCellType());
                        if (rawData != null) {
                            return evaluateCellData(cell, rawData, headerRowData);
                        }
                        return new ExcelCellData(cell.getColumnIndex(), getColumnHeader(cell, headerRowData, cell.getColumnIndex()), null, null, null);
                    })
                    .filter(Objects::nonNull)
                    .collect(Collectors.toList());
        }
    }

    private Object getValue(Cell cell, CellType cellType) {
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue() ? Boolean.TRUE : Boolean.FALSE;
            case FORMULA:
                return (cell.getCachedFormulaResultType() != null)
                        ? getValue(cell, cell.getCachedFormulaResultType())
                        : cell.getCellFormula();
            case ERROR:
                return cell.getErrorCellValue();
            case BLANK:
            case _NONE:
            default:
                return null;
        }
    }

    private ExcelCellData evaluateCellData(Cell cell, Object cellValueString, List<ExcelCellData> headerRowData) {
        final int columnIndex = cell.getColumnIndex();
        final String columnHeader = getColumnHeader(cell, headerRowData, columnIndex);
        switch (cell.getCellType()) {
            case ERROR:
                return new ExcelCellData(columnIndex, columnHeader, cellValueString, "ERROR:" + cellValueString);
            case BOOLEAN:
            case FORMULA:
                return new ExcelCellData(columnIndex, columnHeader, cellValueString, cellValueString);
            case STRING: // We haven't seen this yet.
                var rtsi = new XSSFRichTextString(cellValueString.toString());
                return new ExcelCellData(columnIndex, columnHeader, cellValueString, rtsi.toString());
            case NUMERIC:
                final var formatString = cell.getCellStyle().getDataFormatString();
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new ExcelCellData(columnIndex, columnHeader, cell.getNumericCellValue(), cell.getDateCellValue(), formatString);
                } else {
                    return new ExcelCellData(columnIndex, columnHeader, cellValueString, cellValueString, formatString);
                }
            default:
                return null;
        }
    }

    private String getColumnHeader(Cell cell, List<ExcelCellData> headerRowData, int columnIndex) {
        if (headerRowData == null) {
            return getCellName(cell);
        } else {
            var header = headerRowData.stream().filter(h -> h.getColumnIndex() == columnIndex).findFirst().orElse(null);
            return header == null ? null : header.getFormattedData().toString().trim();
        }
    }

    private String getCellName(Cell cell) {
        return CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1);
    }

    public void openSheet(String sheetName) {
        if (sheetName == null || sheetName.isEmpty()) {
            throw new DataImporterRuntimeException("'" + sheetName + "' cannot be empty");
        }
        if (workbook.getSheet(sheetName) == null) {
            throw new DataImporterRuntimeException("Sheet with a name '" + sheetName + "' not found.");
        }
        this.sheet = workbook.getSheet(sheetName);
    }

    @Override
    public void close() throws Exception {
        if (workbook != null) {
            workbook.close();
        }
        workbook = null;
    }
}
