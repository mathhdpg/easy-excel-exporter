package com.easyexcel.strategies.cellvalue;

import java.sql.Timestamp;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public class SqlTimestampCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        cell.setCellValue((Timestamp) objValue);
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.DATE_DD_MM_YYYY_HH_MM;
    }

}
