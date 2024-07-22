// package com.easyexcel;

// import org.apache.poi.ss.usermodel.*;
// import org.apache.poi.ss.usermodel.DateUtil;
// import org.apache.poi.ss.util.CellRangeAddress;
// import org.apache.poi.xssf.streaming.SXSSFWorkbook;
// import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// import java.io.ByteArrayInputStream;
// import java.io.ByteArrayOutputStream;
// import java.io.InputStream;
// import java.math.BigDecimal;
// import java.math.BigInteger;
// import java.nio.charset.StandardCharsets;
// import java.sql.Date;
// import java.sql.Time;
// import java.sql.Timestamp;
// import java.util.Base64;
// import java.util.Calendar;
// import java.util.List;

// public class ExcelUtil {

// @SuppressWarnings("unused")
// public static byte[] geraXLSXModelo(InputStream modelo, String aba, int
// linha, int coluna, List<Object[]> linhas) {
// return geraXLSXModelo(modelo, aba, linha, coluna, linhas, null);
// }

// @SuppressWarnings("unused")
// public static byte[] geraXLSXModelo(InputStream modelo, String aba, int
// linha, int coluna, List<Object[]> linhas,
// String activeSheet) {
// byte[] retorno;
// Boolean autoSize = true;
// try {
// linha--;
// coluna--;

// Workbook wb = new XSSFWorkbook(modelo);

// Sheet sheet;
// if (aba == null || aba.trim().isEmpty())
// sheet = wb.getSheetAt(0);
// else
// sheet = wb.getSheet(aba);

// int rowCount = linha;
// int colCount = coluna;

// int tamanhoLinha = 1;
// if (linhas.size() > 0)
// tamanhoLinha = linhas.get(0).length;

// if (linhas.size() > 0) {
// CellStyle[] styles = carregaStylesModelo(sheet, tamanhoLinha, rowCount,
// colCount);
// adicionaLinhasPlanilha(sheet, styles, linhas, rowCount, coluna);
// }

// if (autoSize) {
// for (int i = 0; i < tamanhoLinha; i++) {
// sheet.autoSizeColumn(i, true);
// }
// }

// ByteArrayOutputStream baos = new ByteArrayOutputStream();
// wb.setForceFormulaRecalculation(true);
// if (activeSheet != null) {
// wb.setActiveSheet(wb.getSheetIndex(activeSheet));
// }
// wb.write(baos);
// wb.close();

// retorno = baos.toByteArray();
// baos.close();
// return retorno;
// } catch (Exception e) {
// e.printStackTrace();
// return null;
// }
// }

// private static CellStyle[] carregaStylesModelo(Sheet sheet, int tamanhoLinha,
// int rowCount, int colCount) {
// Row rowModelo = sheet.getRow(rowCount);
// CellStyle[] styles = new CellStyle[tamanhoLinha];
// for (int i = 0; i < styles.length; i++) {
// Cell celula = (Utils.isNotEmpty(rowModelo)) ? rowModelo.getCell(colCount + i)
// : null;
// if (celula == null) {
// styles[i] = createCellStyle(sheet.getWorkbook(),
// EnumFormatacaoPadrao.SEM_FORMATACAO.getExcelFormat());
// } else {
// styles[i] = celula.getCellStyle();
// }
// }
// if (Utils.isNotEmpty(rowModelo))
// sheet.removeRow(rowModelo);
// return styles;
// }

// private static int adicionaLinhasPlanilha(Sheet sheet, CellStyle[] styles,
// List<Object[]> linhas, int rowCount,
// int coluna) {
// int colCount;
// for (Object[] lin : linhas) {
// Row row = sheet.createRow(rowCount++);
// colCount = coluna;
// for (int i = 0; i < lin.length; i++) {
// Cell cell = row.createCell(colCount++);
// cell.setCellStyle(styles[i]);
// setCellValue(cell, lin[i]);
// }
// }
// return rowCount;
// }
// }