package com.easyexcel.strategies.cellvalue;

import java.math.BigInteger;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public class BigIntegerCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        cell.setCellValue(((BigInteger) objValue).longValue());
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.INTEGER;
    }
}
