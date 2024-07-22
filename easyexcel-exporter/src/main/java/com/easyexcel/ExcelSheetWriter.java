package com.easyexcel;

import java.util.List;
import java.util.stream.IntStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.easyexcel.beans.ExcelCell;
import com.easyexcel.beans.ExcelFile;
import com.easyexcel.beans.ExcelRow;
import com.easyexcel.beans.ExcelSheet;
import com.easyexcel.beans.SimpleExcelFile;
import com.easyexcel.enums.CustomCellFormat;
import com.easyexcel.strategies.CellValueContext;
import com.easyexcel.utils.ExcelCellFactory;
import com.easyexcel.utils.ExcelCellStyleHelper;

public class ExcelSheetWriter {

    private static final int WIDTH_ARROW_BUTTON = 700;

    private final SXSSFWorkbook workbook;

    public ExcelSheetWriter() {
        this.workbook = new SXSSFWorkbook(1000);
    }

    public Workbook create(ExcelFile excelFile) {
        excelFile.getSheets().forEach(this::createSheet);
        return workbook;
    }

    public Workbook create(SimpleExcelFile simpleExcelFile) {
        String sheetName = simpleExcelFile.getSheetName();
        List<String> header = simpleExcelFile.getHeader();
        List<Object[]> lines = simpleExcelFile.getLines();
        List<CustomCellFormat> formatters = simpleExcelFile.getFormatters();

        SXSSFSheet sheet = workbook.createSheet(sheetName == null || sheetName.trim().isEmpty() ? "Plan1" : sheetName);
        Row headerRow = sheet.createRow(0);
        IntStream.range(0, header.size()).forEach(cellIndex -> {
            ExcelCell headerCell = ExcelCellFactory.createDefaultHeaderCell(header.get(cellIndex));
            createCell(headerRow, headerCell, cellIndex);
        });
        IntStream.range(0, lines.size()).forEach(rowIndex -> {
            Object[] line = lines.get(rowIndex);
            Row row = sheet.createRow(rowIndex + 1);
            IntStream.range(0, line.length).forEach(cellIndex -> {
                CustomCellFormat format = formatters != null ? formatters.get(cellIndex) : null;
                ExcelCell lineCell = ExcelCellFactory.createDefaultCell(line[cellIndex], format);
                createCell(row, lineCell, cellIndex);
            });
        });
        autoSizeAndFilter(sheet, header.size() - 1);
        return workbook;
    }

    public void createSheet(ExcelSheet sheetTemplate) {
        SXSSFSheet sheet = workbook.createSheet(sheetTemplate.getName());
        createRows(sheet, sheetTemplate.getRows());
        autoSizeAndFilter(sheet, sheetTemplate.getRows().get(0).getCells().size() - 1);
    }

    private void autoSizeAndFilter(SXSSFSheet sheet, Integer columnsLength) {
        workbook.setForceFormulaRecalculation(true);
        sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, columnsLength));
        sheet.trackAllColumnsForAutoSizing();
        autoSizeColumns(sheet, 0, columnsLength);
    }

    private void createRows(SXSSFSheet sheet, List<ExcelRow> rows) {
        IntStream.range(0, rows.size()).forEach(rowIndex -> {
            Row row = sheet.createRow(rowIndex);
            ExcelRow rowTemplate = rows.get(rowIndex);
            List<ExcelCell> cells = rowTemplate.getCells();
            IntStream.range(0, cells.size()).forEach(cellIndex -> createCell(row, cells.get(cellIndex), cellIndex));
        });
    }

    private void createCell(Row row, ExcelCell cellDetails, int cellIndex) {
        Cell cell = row.createCell(cellIndex);
        CellValueContext.setCellValue(cell, cellDetails.getValue());
        cell.setCellStyle(new ExcelCellStyleHelper(workbook, cellDetails).getCellStyle());
    }

    public void autoSizeColumns(Sheet sheet, int fromColumn, int toColumn) {
        for (int i = fromColumn; i <= toColumn; i++) {
            sheet.autoSizeColumn(i);
            try {
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + WIDTH_ARROW_BUTTON);
            } catch (Exception e) {
                // don't do anything - just let autosize handle it
            }
        }
    }
}
