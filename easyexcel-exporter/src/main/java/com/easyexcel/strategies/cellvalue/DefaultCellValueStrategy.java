package com.easyexcel.strategies.cellvalue;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public class DefaultCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        if (objValue == null || objValue.toString().isEmpty()) {
            cell.setCellValue("");

        } else {
            cell.setCellValue(objValue.toString());
        }
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.TEXT;
    }

}
