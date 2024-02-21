package org.jtgm.core;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class DefaultExcelExtractor implements ExcelExtractor{
    @Override
    public void extract(String filePath) {
        try {
            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);

            Sheet sheet = workbook.getSheetAt(0);

            Row rowI2 = sheet.getRow(1);
            Cell cellI2 = rowI2.getCell(8);

            Row rowK2 = sheet.getRow(1);
            Cell cellK2 = rowK2.createCell(10);
            cellK2.setCellValue(cellI2.getStringCellValue());

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
                System.out.println("Value copied from I2 to K2.");
            } catch (IOException e) {
                System.err.println("Error writing to the file: " + e.getMessage());
            }
        } catch (Exception e) {
           System.err.println("Error occur");
        }
    }
}
