package com.easyexcel.strategies.cellvalue;

import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public class BigDecimalCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        cell.setCellValue(((BigDecimal) objValue).doubleValue());
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.DECIMAL;
    }

}
