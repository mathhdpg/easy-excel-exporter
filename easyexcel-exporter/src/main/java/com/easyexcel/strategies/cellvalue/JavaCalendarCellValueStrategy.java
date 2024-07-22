package com.easyexcel.strategies.cellvalue;

import java.util.Calendar;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public class JavaCalendarCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        cell.setCellValue((Calendar) objValue);
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.DATE_DD_MM_YYYY;
    }

}
