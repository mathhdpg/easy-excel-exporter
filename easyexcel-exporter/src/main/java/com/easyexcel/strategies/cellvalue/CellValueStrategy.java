package com.easyexcel.strategies.cellvalue;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;

public interface CellValueStrategy {

    void setCellValue(Cell cell, Object objValue);

    CustomCellFormat getDefaultCellFormat();

}
