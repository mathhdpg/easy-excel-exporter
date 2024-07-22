package com.easyexcel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

import org.apache.poi.ss.usermodel.Workbook;

import com.easyexcel.beans.ExcelFile;
import com.easyexcel.beans.SimpleExcelFile;

public class ExcelExporter {

    public static byte[] generateXLSX(ExcelFile excelFile) {
        try {
            return getByteArray(
                    new ExcelSheetWriter()
                            .create(excelFile));
        } catch (Exception e) {
            e.printStackTrace();
            return "Erro ao gerar o Arquivo em Excel.".getBytes(StandardCharsets.ISO_8859_1);
        }
    }

    public static byte[] generateXLSX(SimpleExcelFile simpleExcelFile) {
        try {
            return getByteArray(
                    new ExcelSheetWriter()
                            .create(
                                    simpleExcelFile.getHeader(),
                                    simpleExcelFile.getLines(),
                                    simpleExcelFile.getFormatters(),
                                    simpleExcelFile.getSheetName()));
        } catch (Exception e) {
            e.printStackTrace();
            return "Erro ao gerar o Arquivo em Excel.".getBytes(StandardCharsets.ISO_8859_1);
        }
    }

    private static byte[] getByteArray(Workbook wb) {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            wb.write(baos);
            return baos.toByteArray();
        } catch (IOException e) {
            e.printStackTrace();
            return "Erro ao gerar o Arquivo em Excel.".getBytes(StandardCharsets.ISO_8859_1);
        }
    }
}
