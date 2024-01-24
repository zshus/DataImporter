package dataimporter.implementation.utils;


import dataimporter.implementation.enums.ExcelExtension;

public class DataImporterUtils {

    private DataImporterUtils() {

    }

    public static ExcelExtension getFileExtension(String fileName) {
        final int lastdot = fileName.lastIndexOf(".");
        if (lastdot < 0) {
            throw new DataImporterRuntimeException("Found file has no extension to derive format from.");
        }
        switch (fileName.substring(lastdot)) {
            case ".xls":
                return ExcelExtension.XLS;
            case ".xlsx":
                return ExcelExtension.XLSX;
            default:
                return ExcelExtension.UNKNOWN;
        }
    }

    public static String sanitizeName(String name) {
        //Applying library conversion logic except reserved keywords
        name = name.replaceAll("[^a-zA-Z0-9_ ]+", "");
        name = name.replaceAll("[\\s\\xa0]+", " ").trim();
        return name.replaceAll("\\W+", "_");
    }
}
