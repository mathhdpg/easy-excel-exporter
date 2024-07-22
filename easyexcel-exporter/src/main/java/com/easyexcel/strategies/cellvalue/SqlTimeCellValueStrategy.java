package com.easyexcel.strategies.cellvalue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import com.easyexcel.enums.CustomCellFormat;

public class SqlTimeCellValueStrategy implements CellValueStrategy {

    @Override
    public void setCellValue(Cell cell, Object objValue) {
        String hora = objValue.toString();
        cell.setCellValue(DateUtil.convertTime(hora));
    }

    @Override
    public CustomCellFormat getDefaultCellFormat() {
        return CustomCellFormat.DATE_HH_MM;
    }

}
