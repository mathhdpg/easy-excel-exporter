package com.easyexcel;

import static com.easyexcel.utils.ExcelCellFactory.createDefaultCell;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import com.easyexcel.beans.ExcelFile;
import com.easyexcel.beans.ExcelRow;
import com.easyexcel.beans.ExcelSheet;
import com.easyexcel.beans.SimpleExcelFile;
import com.easyexcel.enums.CustomCellFormat;
import com.easyexcel.utils.ExcelCellFactory;

public class Example {

    public static void main(String[] args) throws Exception {
        // excelFileExample();
        simpleExcelFileExample();
    }

    private static void simpleExcelFileExample() throws FileNotFoundException, IOException {
        List<String> header = Arrays.asList(
                "Inteiro",
                "Decimal",
                "String",
                "Data",
                "Timestamp",
                "Inteiro positivo",
                "inteiro Negativo",
                "Decimal positivo",
                "Decimal Negativo",
                "Monetario positivo",
                "Monetario Negativo",
                "Monetario positivo colorido",
                "Monetario Negativo colorido",
                "Percentual positivo",
                "Percentual Negativo",
                "Percentual positivo colorido",
                "Percentual Negativo colorido");

        List<Object[]> lines = Arrays.asList(
                new Object[] { 1, 1.0, "string", new Date(), new Timestamp(System.currentTimeMillis()), 1, -1, 1.1,
                        -1.1, 1.2, -1.2, 1.3, -1.3,
                        1.4, -1.4, 1.5, -1.5 },
                new Object[] { 1, 1.0, "string", new Date(), new Timestamp(System.currentTimeMillis()), 1, -1, 1.0,
                        -1.0, 1.0, -1.0, 1.0, -1.0,
                        1.0, -1.0, 1.0, -1.0 });

        List<CustomCellFormat> formatters = Arrays.asList(
                CustomCellFormat.INTEGER,
                CustomCellFormat.DECIMAL,
                CustomCellFormat.TEXT,
                CustomCellFormat.DATE_DD_MM,
                CustomCellFormat.DATE_DD_MM_YYYY_HH_MM_SS,
                CustomCellFormat.INTEGER,
                CustomCellFormat.INTEGER,
                CustomCellFormat.DECIMAL,
                CustomCellFormat.DECIMAL,
                CustomCellFormat.MONETARY_BRL,
                CustomCellFormat.MONETARY_BRL,
                CustomCellFormat.MONETARY_BRL_COLLORED,
                CustomCellFormat.MONETARY_BRL_COLLORED,
                CustomCellFormat.PERCENTAGE,
                CustomCellFormat.PERCENTAGE,
                CustomCellFormat.PERCENTAGE_COLLORED,
                CustomCellFormat.PERCENTAGE_COLLORED);
        SimpleExcelFile simpleExcelFile = new SimpleExcelFile(header, lines, formatters, "Teste Sheet Name");
        byte[] xlsx = ExcelExporter.generateXLSX(simpleExcelFile);
        String caminho = "D:/teste1.xlsx";
        exportarArquivo(xlsx, caminho);
    }

    private static void excelFileExample() throws FileNotFoundException, IOException {
        ExcelRow row1 = new ExcelRow();
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Inteiro"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Decimal"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("String"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Data"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Timestamp"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Inteiro positivo"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("inteiro Negativo"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Decimal positivo"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Decimal Negativo"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Monetario positivo"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Monetario Negativo"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Monetario positivo colorido"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Monetario Negativo colorido"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Percentual positivo"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Percentual Negativo"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Percentual positivo colorido"));
        row1.addCell(ExcelCellFactory.createDefaultHeaderCell("Percentual Negativo colorido"));

        ExcelRow row2 = new ExcelRow();
        row2.addCell(createDefaultCell(1));
        row2.addCell(createDefaultCell(1.5));
        row2.addCell(createDefaultCell("String"));
        row2.addCell(createDefaultCell(new Date()));
        row2.addCell(createDefaultCell(new Timestamp(System.currentTimeMillis())));
        row2.addCell(createDefaultCell(1, CustomCellFormat.INTEGER));
        row2.addCell(createDefaultCell(-1, CustomCellFormat.INTEGER));
        row2.addCell(createDefaultCell(1.5, CustomCellFormat.DECIMAL));
        row2.addCell(createDefaultCell(-1.5, CustomCellFormat.DECIMAL));
        row2.addCell(createDefaultCell(new BigDecimal(1.5), CustomCellFormat.MONETARY_BRL));
        row2.addCell(createDefaultCell(new BigDecimal(-1.5), CustomCellFormat.MONETARY_BRL));
        row2.addCell(createDefaultCell(new BigDecimal(1.5), CustomCellFormat.MONETARY_BRL_COLLORED));
        row2.addCell(createDefaultCell(new BigDecimal(-1.5), CustomCellFormat.MONETARY_BRL_COLLORED));
        row2.addCell(createDefaultCell(new BigDecimal(1.5), CustomCellFormat.PERCENTAGE));
        row2.addCell(createDefaultCell(new BigDecimal(-1.5), CustomCellFormat.PERCENTAGE));
        row2.addCell(createDefaultCell(new BigDecimal(1.5), CustomCellFormat.PERCENTAGE_COLLORED));
        row2.addCell(createDefaultCell(new BigDecimal(-1.5), CustomCellFormat.PERCENTAGE_COLLORED));

        ExcelSheet sheet = new ExcelSheet();
        sheet.setName("teste");
        sheet.addRow(row1);
        sheet.addRow(row2);

        ExcelFile excelFile = new ExcelFile();
        excelFile.setName("teste");
        excelFile.addSheet(sheet);

        byte[] bytes = ExcelExporter.generateXLSX(excelFile);

        String caminho = "D:/" + excelFile.getName() + ".xlsx";
        exportarArquivo(bytes, caminho);
    }

    private static void exportarArquivo(byte[] bytes, String caminho) throws FileNotFoundException, IOException {
        File file = new File(caminho);
        FileOutputStream in = new FileOutputStream(file);
        in.write(bytes);
        in.close();
    }

}
