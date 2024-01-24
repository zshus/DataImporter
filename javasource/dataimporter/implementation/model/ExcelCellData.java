package dataimporter.implementation.model;

import java.util.Objects;

public class ExcelCellData {
    private final int columnIndex;
    private final String columnHeader;
    private final Object rawData;
    private final String displayMask;
    private final Object formattedData;

    public ExcelCellData(int columnIndex, String columnHeader, Object rawData, Object formattedData) {
        this(columnIndex, columnHeader, rawData, formattedData, null);
    }

    public ExcelCellData(int columnIndex, String columnHeader, Object rawData, Object formattedData, String displayMask) {
        this.columnIndex = columnIndex;
        this.columnHeader = columnHeader;
        this.rawData = rawData;
        this.formattedData = formattedData;
        this.displayMask = displayMask;
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public String getColumnHeader() {
        return columnHeader;
    }

    public Object getRawData() {
        return rawData;
    }

    public String getDisplayMask() {
        return displayMask;
    }

    public Object getFormattedData() {
        return formattedData;
    }

    @Override
    public int hashCode() {
        return columnIndex + 31 * Objects.hash(rawData, displayMask, formattedData);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        ExcelCellData that = (ExcelCellData) o;
        return columnIndex == that.columnIndex &&
                columnHeader == that.columnHeader &&
                Objects.equals(rawData, that.rawData) &&
                Objects.equals(displayMask, that.displayMask) &&
                Objects.equals(formattedData, that.formattedData);
    }

    @Override
    public String toString() {
        return "ExcelCellData{ " +
                "colNo=" + columnIndex +
                ", colName=" + columnHeader +
                ", rawData=" + rawData +
                ", formattedData=" + formattedData +
                ", displayMask='" + displayMask + '\'' +
                " }";
    }
}