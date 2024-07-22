package com.easyexcel.strategies;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;

import com.easyexcel.enums.CustomCellFormat;
import com.easyexcel.strategies.cellvalue.BigDecimalCellValueStrategy;
import com.easyexcel.strategies.cellvalue.BigIntegerCellValueStrategy;
import com.easyexcel.strategies.cellvalue.CellValueStrategy;
import com.easyexcel.strategies.cellvalue.DefaultCellValueStrategy;
import com.easyexcel.strategies.cellvalue.DoubleCellValueStrategy;
import com.easyexcel.strategies.cellvalue.IntegerCellValueStrategy;
import com.easyexcel.strategies.cellvalue.JavaCalendarCellValueStrategy;
import com.easyexcel.strategies.cellvalue.JavaDateCellValueStrategy;
import com.easyexcel.strategies.cellvalue.LongCellValueStrategy;
import com.easyexcel.strategies.cellvalue.SqlDateCellValueStrategy;
import com.easyexcel.strategies.cellvalue.SqlTimeCellValueStrategy;
import com.easyexcel.strategies.cellvalue.SqlTimestampCellValueStrategy;

public class CellValueContext {

    private static final Map<Class<?>, CellValueStrategy> strategies = new HashMap<>();

    static {
        strategies.put(BigDecimal.class, new BigDecimalCellValueStrategy());
        strategies.put(BigInteger.class, new BigIntegerCellValueStrategy());
        strategies.put(Double.class, new DoubleCellValueStrategy());
        strategies.put(Integer.class, new IntegerCellValueStrategy());
        strategies.put(Calendar.class, new JavaCalendarCellValueStrategy());
        strategies.put(java.util.Date.class, new JavaDateCellValueStrategy());
        strategies.put(Long.class, new LongCellValueStrategy());
        strategies.put(java.sql.Date.class, new SqlDateCellValueStrategy());
        strategies.put(Timestamp.class, new SqlTimestampCellValueStrategy());
        strategies.put(Time.class, new SqlTimeCellValueStrategy());
    }

    private static CellValueStrategy getStrategy(Object objValue) {
        CellValueStrategy strategy = strategies.getOrDefault(
                objValue != null ? objValue.getClass() : Object.class,
                new DefaultCellValueStrategy());
        return strategy;
    }

    public static void setCellValue(Cell cell, Object objValue) {
        getStrategy(objValue).setCellValue(cell, objValue);
    }

    public static CustomCellFormat getDefaultCellFormat(Object objValue) {
        return getStrategy(objValue).getDefaultCellFormat();
    }

}