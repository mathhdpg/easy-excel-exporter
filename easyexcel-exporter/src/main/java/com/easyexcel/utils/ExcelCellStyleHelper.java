package com.easyexcel.utils;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import com.easyexcel.beans.ExcelCell;
import com.easyexcel.enums.CustomCellFormat;
import com.easyexcel.strategies.CellValueContext;

public class ExcelCellStyleHelper {

    private Workbook workbook;

    private ExcelCell cell;

    public ExcelCellStyleHelper(Workbook workbook, ExcelCell cell) {
        this.workbook = workbook;
        this.cell = cell;
    }

    public CellStyle getCellStyle() {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(getCellFontConfig());
        cellStyle.setBorderTop(cell.getBorderTop());
        cellStyle.setBorderBottom(cell.getBorderBotton());
        cellStyle.setBorderLeft(cell.getBorderLeft());
        cellStyle.setBorderRight(cell.getBorderRight());
        setCellFormat(cellStyle);
        return cellStyle;
    }

    private void setCellFormat(CellStyle cellStyle) {
        Object value = cell.getValue();
        CustomCellFormat cellFormat = cell.getCellFormat();
        if (cellFormat == null) {
            cellFormat = CellValueContext.getDefaultCellFormat(value);
        }
        cellStyle.setDataFormat(this.workbook.createDataFormat().getFormat(cellFormat.getValue()));
    }

    private Font getCellFontConfig() {
        Font font = this.workbook.createFont();
        font.setFontName(cell.getFontName());
        font.setBold(cell.getIsBold());
        font.setItalic(cell.isItalic());
        font.setFontHeightInPoints((short) cell.getFontSize());
        font.setColor(cell.getFontColor().getIndex());
        return font;
    }
}