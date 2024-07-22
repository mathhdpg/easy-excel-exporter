package com.easyexcel.beans;

import java.util.ArrayList;
import java.util.List;

public class ExcelSheet {

    private String name;
    private List<ExcelRow> rows;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<ExcelRow> getRows() {
        return rows;
    }

    public void setRows(List<ExcelRow> rows) {
        this.rows = rows;
    }

    public void addRow(ExcelRow row) {
        if (rows == null) {
            rows = new ArrayList<ExcelRow>();
        }
        rows.add(row);
    }
}
