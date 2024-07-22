package com.easyexcel.beans;

import java.util.ArrayList;
import java.util.List;

public class ExcelFile {

    private String name;
    private List<ExcelSheet> sheets;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<ExcelSheet> getSheets() {
        return sheets;
    }

    public void setSheets(List<ExcelSheet> sheets) {
        this.sheets = sheets;
    }

    public void addSheet(ExcelSheet sheet) {
        if (sheets == null) {
            sheets = new ArrayList<ExcelSheet>();
        }
        sheets.add(sheet);
    }

}
