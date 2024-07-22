package com.easyexcel.enums;

public enum CustomCellFormat {
    TEXT("text"),

    INTEGER("#,##0"),
    INTEGER_COLLORED("#,##0;[Red]-#,##0"),

    DECIMAL("#,##0.00"),
    DECIMAL_COLLORED("#,##0.00;[Red]-#,##0.00"),

    PERCENTAGE("#,##0.00%"),
    PERCENTAGE_COLLORED("#,##0.00%;[Red]-#,##0.00%"),

    MONETARY_BRL("R$ #,##0.00"),
    MONETARY_BRL_COLLORED("R$ #,##0.00;[Red]-R$ #,##0.00"),

    MONETARY_USD("$ #,##0.00"),
    MONETARY_USD_COLLORED("$ #,##0.00;[Red]-$ #,##0.00"),

    DATE_DD_MM("dd/mm"),
    DATE_DD_MM_YY("dd/mm/yy"),
    DATE_DD_MM_YYYY("dd/mm/yyyy"),

    DATE_MM_YY("mm/yy"),
    DATE_MM_YYYY("mm/yyyy"),

    DATE_DD_MM_HH_MM("dd/mm hh:mm"),
    DATE_DD_MM_YY_HH_MM("dd/mm/yy hh:mm"),
    DATE_DD_MM_YYYY_HH_MM("dd/mm/yyyy hh:mm"),

    DATE_HH_MM("HH:mm"),

    DATE_DD_MM_YY_HH_MM_SS("dd/mm/yy hh:mm:ss"),
    DATE_DD_MM_YYYY_HH_MM_SS("dd/mm/yyyy hh:mm:ss"),
    DATE_HH_MM_SS("HH:mm:ss"),
    ;

    private String value;

    CustomCellFormat(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }

}