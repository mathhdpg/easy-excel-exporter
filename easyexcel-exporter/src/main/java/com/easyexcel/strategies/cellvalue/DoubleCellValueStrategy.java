package com.easyexcel.strategies.cellvalue;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public class DoubleCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        cell.setCellValue((Double) objValue);
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.DECIMAL;
    }
}
