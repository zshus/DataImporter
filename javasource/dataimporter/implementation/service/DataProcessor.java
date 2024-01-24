package dataimporter.implementation.service;


import com.mendix.core.Core;
import com.mendix.core.CoreException;
import com.mendix.logging.ILogNode;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import com.mendix.systemwideinterfaces.core.meta.IMetaPrimitive;
import dataimporter.implementation.model.ExcelCellData;
import dataimporter.implementation.utils.DataImporterRuntimeException;
import dataimporter.implementation.utils.DataImporterUtils;
import dataimporter.implementation.utils.DataReaderException;
import dataimporter.proxies.ColumnAttributeMapping;
import dataimporter.proxies.Sheet;
import dataimporter.proxies.constants.Constants;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.util.RecordFormatException;

import java.io.File;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.util.*;
import java.util.stream.Collectors;

public class DataProcessor {

    public static final ILogNode logNode = Core.getLogger(Constants.getLogNode());
    static final String ERROR_WHILE_IMPORTING = "Error while importing: '";
    static final String MS_BECAUSE = " ms, because: ";
    static final String STARTED = " started.";
    static final String FROM_SHEET = " from sheet ";
    private static String sheetName;

    private DataProcessor() {
    }

    public static void startImport(IContext context, IMendixObject mappingTemplate, File excelFile, String excelFileName, List<IMendixObject> importedList) throws DataImporterRuntimeException, CoreException {
        Map<Sheet, List<ColumnAttributeMapping>> sheetColumnMappingMap = new HashMap<>();
        List<IMendixObject> templateSheets = Core.retrieveByPath(context, mappingTemplate, Sheet.MemberNames.Sheet_Template.toString());
        for (IMendixObject templateSheetObject : templateSheets) {
            List<ColumnAttributeMapping> columnAttributeMappings = new ArrayList<>();
            List<IMendixObject> columnAttributeMappingObjects = Core.retrieveByPath(context, templateSheetObject, ColumnAttributeMapping.MemberNames.ColumnAttributeMapping_Sheet.toString());
            for (IMendixObject columnAttributeMapping : columnAttributeMappingObjects) {
                columnAttributeMappings.add(ColumnAttributeMapping.initialize(context, columnAttributeMapping));
            }
            sheetColumnMappingMap.put(Sheet.initialize(context, templateSheetObject), columnAttributeMappings);
        }
        var importStartTime = 0L;
        try {
            importStartTime = System.nanoTime();
            switch (DataImporterUtils.getFileExtension(excelFileName)) {
                case XLS:
                case XLSX:
                    for (Map.Entry<Sheet, List<ColumnAttributeMapping>> entry : sheetColumnMappingMap.entrySet()) {
                        parseData(context, excelFile, entry.getKey(), entry.getValue(), importedList);
                    }
                    break;
                case UNKNOWN:
                    throw new CoreException("File extension is not an Excel extension ('.xls' or '.xlsx').");
                default:
                    throw new CoreException("File extension is not an Excel extension ('.xls' or '.xlsx').");
            }
            logNode.info("Successfully finished importing '" + importedList.size() + "' rows of '" + sheetName + "' sheet from excelFile: '" + excelFileName + "' in '" + ((System.nanoTime() - importStartTime) / 1000000) + " ms'");
        } catch (OLE2NotOfficeXmlFileException e) {
            logNode.error(ERROR_WHILE_IMPORTING + excelFileName + "' " + ((System.nanoTime() - importStartTime) / 1000000) + MS_BECAUSE + e.getMessage());
            throw new DataImporterRuntimeException("Document could not be imported because this excelFile is an XLS and not an XLSX excelFile. Please make sure the excelFile is valid and has the correct extension.");
        } catch (NotOfficeXmlFileException e) {
            logNode.error(ERROR_WHILE_IMPORTING + excelFileName + "' " + ((System.nanoTime() - importStartTime) / 1000000) + MS_BECAUSE + e.getMessage());
            throw new DataImporterRuntimeException("Document could not be imported because this excelFile is not XLS or XLSX. Please make sure the excelFile is valid and has the correct extension.");
        } catch (RecordFormatException e) {
            logNode.error(ERROR_WHILE_IMPORTING + excelFileName + "' " + ((System.nanoTime() - importStartTime) / 1000000) + MS_BECAUSE + e.getMessage());
            throw new DataImporterRuntimeException("Document could not be imported because one of its cell values is invalid or cannot be read.");
        } catch (EncryptedDocumentException e) {
            logNode.error(ERROR_WHILE_IMPORTING + excelFileName + "' " + ((System.nanoTime() - importStartTime) / 1000000) + MS_BECAUSE + e.getMessage());
            throw new DataImporterRuntimeException("Document could not be imported because it is encrypted.");
        } catch (Exception e) {
            logNode.error(ERROR_WHILE_IMPORTING + excelFileName + "' " + ((System.nanoTime() - importStartTime) / 1000000) + MS_BECAUSE + e.getMessage());
            throw new CoreException("Uploaded excel file could not be imported, because: " + e.getMessage(), e);
        } finally {
            if (excelFile != null) {
                try {
                    Files.delete(excelFile.toPath());
                } catch (final Exception ignored) {
                    logNode.error("Could not delete temp excelFile.");
                }
            }
        }
    }

    public static void parseData(IContext context, File file, Sheet sheetMendixObject, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects, List<IMendixObject> importedList) {
        try (var dataReader = new DataReader(file)) {
            sheetName = sheetMendixObject.getSheetName();
            dataReader.openSheet(sheetName);
            if (logNode.isTraceEnabled()) {
                logNode.trace("Reading excel header row from sheet: '" + sheetName + "'" + STARTED);
            }
            List<ExcelCellData> headerRowData = dataReader.readHeaderRow(sheetMendixObject.getHeaderRowStartsAt() - 1);
            if (logNode.isTraceEnabled()) {
                logNode.trace("Reading excel header row from sheet: '" + sheetName + "' finished. Found '" + headerRowData.size() + "' columns.");
            }
            if (headerRowData == null || headerRowData.isEmpty()) {
                throw new DataImporterRuntimeException("No column information could be found in sheet: '" + sheetName + "'");
            }

            Set<String> headerColumnNames = headerRowData.stream()
                    .map(ExcelCellData::getFormattedData)
                    .map(Object::toString)
                    .collect(Collectors.toSet());
            for (ColumnAttributeMapping columnAttributeMapping : columnAttributeMappingMendixObjects) {
                if (!headerColumnNames.contains(columnAttributeMapping.getColumnName())) {
                    throw new DataImporterRuntimeException("column with a name: '" + columnAttributeMapping.getColumnName() + "' is not found in sheet: '" + sheetName + "'");
                }
            }

            int dataRowNo = sheetMendixObject.getDataRowStartsAt() - 1;
            while (dataReader.hasNextRow(dataRowNo)) {
                dataRowNo = readExcelRow(context, columnAttributeMappingMendixObjects, dataReader, headerRowData, dataRowNo, importedList);
            }
        } catch (Exception e) {
            throw new DataImporterRuntimeException(e.getMessage(), e);
        }
    }

    private static int readExcelRow(IContext context, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects, DataReader dataReader, List<ExcelCellData> headerRowData, int dataRowNo, List<IMendixObject> importedList) throws DataReaderException {
        try {
            if (logNode.isTraceEnabled()) {
                logNode.trace("Reading excel row: " + dataRowNo + FROM_SHEET + sheetName + STARTED);
            }
            List<ExcelCellData> dataRow = dataReader.readDataRow(dataRowNo, headerRowData);
            if (logNode.isTraceEnabled()) {
                logNode.trace("Reading excel row: " + dataRowNo + FROM_SHEET + sheetName + " finished. Found " + dataRow.size() + " cells.");
            }
            //rows with all empty cells will not be imported
            if (!dataRow.isEmpty()) {
                if (logNode.isTraceEnabled()) {
                    logNode.trace("Importing excel row: " + dataRowNo + FROM_SHEET + sheetName + STARTED);
                }
                importedList.add(processRowData(context, dataRow, columnAttributeMappingMendixObjects));
                if (logNode.isTraceEnabled()) {
                    logNode.trace("Importing excel row: " + dataRowNo + FROM_SHEET + sheetName + " finished.");
                }
            }
            dataRowNo++;
        } catch (Exception e) {
            throw new DataReaderException("Unable to import sheet row '" + dataRowNo + "'" + FROM_SHEET + " '" + sheetName + "'", e);
        }
        return dataRowNo;
    }

    public static IMendixObject processRowData(IContext context, List<ExcelCellData> dataRow, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects) {
        // Store MetaPrimitives in a map to avoid multiple calls to Core API functions
        Map<String, IMetaPrimitive> metaPrimitiveMap = new HashMap<>();
        for (ColumnAttributeMapping attributeMapping : columnAttributeMappingMendixObjects) {
            String attributeName = Core.getMetaPrimitive(attributeMapping.getAttribute()).getName();
            var iMetaPrimitive = Core.getMetaPrimitive(attributeMapping.getAttribute());
            metaPrimitiveMap.put(attributeName, iMetaPrimitive);
        }
        // Create a HashMap to store the ExcelCellData objects for each column header
        Map<String, ExcelCellData> cellDataMap = new HashMap<>();
        for (ExcelCellData excelCellData : dataRow) {
            cellDataMap.put(excelCellData.getColumnHeader(), excelCellData);
        }
        // Create the entity object
        IMendixObject entityObject = Core.instantiate(context, Core.getMetaPrimitive(columnAttributeMappingMendixObjects.get(0).getAttribute()).getParent().getName());
        for (ColumnAttributeMapping attributeMapping : columnAttributeMappingMendixObjects) {
            String attributeName = Core.getMetaPrimitive(attributeMapping.getAttribute()).getName();
            var excelCellData = cellDataMap.get(attributeMapping.getColumnName());
            if (excelCellData != null) {
                var iMetaPrimitive = metaPrimitiveMap.get(attributeName);
                entityObject.setValue(context, attributeName, getMendixTypeObject(iMetaPrimitive, excelCellData));
            }
        }
        return entityObject;
    }

    public static Object getMendixTypeObject(IMetaPrimitive metaPrimitive, ExcelCellData excelCellData) {
        if (excelCellData.getFormattedData() == null) {
            return null;
        }
        logNode.trace("Excel cell is type of: " + excelCellData.getFormattedData().getClass() + " & PrimitiveType is: " + metaPrimitive.getType());
        switch (metaPrimitive.getType()) {
            case String:
            case Boolean:
            case DateTime:
                return excelCellData.getFormattedData();
            case Decimal:
                return new BigDecimal(excelCellData.getFormattedData().toString());
            default:
                return new DataReaderException("Mismatched data type found between excel cell and entity attribute.");
        }
    }
}
