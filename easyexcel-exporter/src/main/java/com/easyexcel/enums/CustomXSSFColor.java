package com.easyexcel.enums;

import org.apache.poi.xssf.usermodel.XSSFColor;

public enum CustomXSSFColor {
    BLACK("#000000"),
    WHITE("#FFFFFF"),
    RED("#FF0000"),
    GREEN("#00FF00"),
    BLUE("#0000FF"),
    YELLOW("#FFFF00"),
    ORANGE("#FFA500"),
    GRAY("#808080"),
    PINK("#FFC0CB"),
    PURPLE("#800080");

    private final String hex;

    CustomXSSFColor(String hex) {
        this.hex = hex;
    }

    public XSSFColor getColor() {
        return new XSSFColor(java.awt.Color.decode(hex), null);
    }
}