package org.jtgm.core.service.impl;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jtgm.core.service.ExcelExtractor;
import org.jtgm.core.util.ExcelUtil;
import org.jtgm.core.exception.GenericErrorException;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@RequiredArgsConstructor
public class DefaultExcelExtractor implements ExcelExtractor {
    private final ExcelUtil excelUtil;

    @Override
    public void extract(MultipartFile file) {
        try {
            String mgroupName = getMgroupName(file);

            Workbook reqWorkbook = new XSSFWorkbook(file.getInputStream());
            Sheet sheet = reqWorkbook.getSheetAt(0);

            excelUtil.execute(sheet, mgroupName);

            Files.createDirectories(Paths.get(System.getProperty("user.home") + "/Processed"));
            String filePath = System.getProperty("user.home") + "/Processed/" + file.getOriginalFilename();
            file.transferTo(new File(filePath));
        }catch (Exception e) {
            e.printStackTrace();
            throw new GenericErrorException("Unable to process file", e);
        }
    }

    private String getMgroupName(MultipartFile file) {
        String name  = file.getOriginalFilename().replaceAll("[\\[\\](){}]",";");
        return name.split(";")[0];
    }

}
