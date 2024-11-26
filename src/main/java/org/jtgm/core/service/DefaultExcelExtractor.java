package org.jtgm.core.service;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jtgm.core.util.ExcelUtil;
import org.jtgm.core.exception.GenericErrorException;
import org.springframework.web.multipart.MultipartFile;

@RequiredArgsConstructor
public class DefaultExcelExtractor implements ExcelExtractor{
    private final ExcelUtil excelUtil;

    @Override
    public void extract(MultipartFile file) {
        try {
            String mgroupName = getMgroupName(file);

            Workbook reqWorkbook = new XSSFWorkbook(file.getInputStream());
            Sheet sheet = reqWorkbook.getSheetAt(0);

            excelUtil.generateWorkBook(sheet, mgroupName);
        }catch (Exception e) {
            e.printStackTrace();
            throw new GenericErrorException("Error occur", e);
        }
    }

    private String getMgroupName(MultipartFile file) {
        String name  = file.getOriginalFilename().replaceAll("[\\[\\](){}]",";");
        return name.split(";")[0];
    }

}
