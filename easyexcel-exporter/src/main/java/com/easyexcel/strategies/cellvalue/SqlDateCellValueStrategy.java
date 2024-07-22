package com.easyexcel.strategies.cellvalue;

import java.sql.Date;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public class SqlDateCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        cell.setCellValue((Date) objValue);
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.DATE_DD_MM_YYYY;
    }

}
