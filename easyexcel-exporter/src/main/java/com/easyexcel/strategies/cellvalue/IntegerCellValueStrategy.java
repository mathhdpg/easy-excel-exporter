package com.easyexcel.strategies.cellvalue;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public class IntegerCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        cell.setCellValue((Integer) objValue);
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.INTEGER;
    }
}
