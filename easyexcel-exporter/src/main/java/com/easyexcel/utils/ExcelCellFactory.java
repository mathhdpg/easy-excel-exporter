package com.easyexcel.utils;

import org.apache.poi.ss.usermodel.BorderStyle;

import com.easyexcel.beans.ExcelCell;
import com.easyexcel.enums.CustomCellFormat;

public class ExcelCellFactory {

    public static ExcelCell createDefaultHeaderCell(String value) {
        ExcelCell headerCell = new ExcelCell(value, CustomCellFormat.TEXT);
        headerCell.setIsBold(true);
        headerCell.setBorderTop(BorderStyle.MEDIUM);
        headerCell.setBorderBotton(BorderStyle.MEDIUM);
        headerCell.setBorderLeft(BorderStyle.MEDIUM);
        headerCell.setBorderRight(BorderStyle.MEDIUM);
        return headerCell;
    }

    public static ExcelCell createDefaultCell(Object value) {
        return createDefaultCell(value, null);
    }

    public static ExcelCell createDefaultCell(Object value, CustomCellFormat cellFormat) {
        ExcelCell cell = new ExcelCell(value, CustomCellFormat.TEXT);
        cell.setBorderTop(BorderStyle.THIN);
        cell.setBorderBotton(BorderStyle.THIN);
        cell.setBorderLeft(BorderStyle.THIN);
        cell.setBorderRight(BorderStyle.THIN);
        cell.setCellFormat(cellFormat);
        return cell;
    }

}
