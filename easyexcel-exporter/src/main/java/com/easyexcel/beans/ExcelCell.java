package com.easyexcel.beans;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import com.easyexcel.enums.CustomCellFormat;
import com.easyexcel.enums.CustomXSSFColor;

public class ExcelCell {
    private Object value;
    private CustomCellFormat cellFormat;

    private BorderStyle borderTop = BorderStyle.THIN;
    private BorderStyle borderBotton = BorderStyle.THIN;
    private BorderStyle borderLeft = BorderStyle.THIN;
    private BorderStyle borderRight = BorderStyle.THIN;

    private String fontName = "Arial";
    private int fontSize = 12;
    private Boolean isBold = Boolean.FALSE;
    private boolean isItalic = Boolean.FALSE;
    private XSSFColor fontColor = CustomXSSFColor.BLACK.getColor();

    public ExcelCell(Object value) {
        this.value = value;
    }

    public ExcelCell(Object value, CustomCellFormat cellFormat) {
        this.value = value;
        this.cellFormat = cellFormat;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public CustomCellFormat getCellFormat() {
        return cellFormat;
    }

    public void setCellFormat(CustomCellFormat cellFormat) {
        this.cellFormat = cellFormat;
    }

    public BorderStyle getBorderTop() {
        return borderTop;
    }

    public void setBorderTop(BorderStyle borderTop) {
        this.borderTop = borderTop;
    }

    public BorderStyle getBorderBotton() {
        return borderBotton;
    }

    public void setBorderBotton(BorderStyle borderBotton) {
        this.borderBotton = borderBotton;
    }

    public BorderStyle getBorderLeft() {
        return borderLeft;
    }

    public void setBorderLeft(BorderStyle borderLeft) {
        this.borderLeft = borderLeft;
    }

    public BorderStyle getBorderRight() {
        return borderRight;
    }

    public void setBorderRight(BorderStyle borderRight) {
        this.borderRight = borderRight;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public Boolean getIsBold() {
        return isBold;
    }

    public void setIsBold(Boolean isBold) {
        this.isBold = isBold;
    }

    public boolean isItalic() {
        return isItalic;
    }

    public void setItalic(boolean isItalic) {
        this.isItalic = isItalic;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public XSSFColor getFontColor() {
        return fontColor;
    }

    public void setFontColor(XSSFColor fontColor) {
        this.fontColor = fontColor;
    }

    public void setFontColor(String hexColor) {
        this.fontColor = new XSSFColor(java.awt.Color.decode(hexColor), null);
    }

    public void setFontColor(CustomXSSFColor customXSSFColor) {
        this.fontColor = customXSSFColor.getColor();
    }
}