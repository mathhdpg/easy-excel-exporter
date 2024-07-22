package com.easyexcel.beans;

import java.util.ArrayList;
import java.util.List;

public class ExcelRow {

    private List<ExcelCell> cells;

    public List<ExcelCell> getCells() {
        return cells;
    }

    public void setCells(List<ExcelCell> cells) {
        this.cells = cells;
    }

    public void addCell(ExcelCell cell) {
        if (cells == null) {
            cells = new ArrayList<ExcelCell>();
        }
        cells.add(cell);
    }
}