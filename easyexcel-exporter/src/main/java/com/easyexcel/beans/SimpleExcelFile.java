package com.easyexcel.beans;

import java.util.List;

import com.easyexcel.enums.CustomCellFormat;

public class SimpleExcelFile {

    private List<String> header;
    private List<Object[]> lines;
    private List<CustomCellFormat> formatters;
    private String sheetName;

    public SimpleExcelFile(List<String> header, List<Object[]> lines, String sheetName) {
        this.header = header;
        this.lines = lines;
        this.sheetName = sheetName;
        this.formatters = null;
    }

    public SimpleExcelFile(List<String> header, List<Object[]> lines, List<CustomCellFormat> formatters,
            String sheetName) {
        this.header = header;
        this.lines = lines;
        this.formatters = formatters;
        this.sheetName = sheetName;
    }

    public List<String> getHeader() {
        return header;
    }

    public List<Object[]> getLines() {
        return lines;
    }

    public List<CustomCellFormat> getFormatters() {
        return formatters;
    }

    public String getSheetName() {
        return sheetName;
    }
}
