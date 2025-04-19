package org.jtgm.core.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jtgm.core.service.ExcelExtractor;
import org.jtgm.core.util.ExcelUtil;
import org.jtgm.core.exception.GenericErrorException;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Paths;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

@RequiredArgsConstructor
@Slf4j
public class DefaultExcelExtractor implements ExcelExtractor {
    private final ExcelUtil excelUtil;

    @Override
    public void extract(File fileRaw) {
        try {
            ZipSecureFile.setMinInflateRatio(0);
            InputStream fileToProcess = new FileInputStream(fileRaw);
            Workbook reqWorkbook = new XSSFWorkbook(fileToProcess);
            fileToProcess.close();

            String mgroupName = getMgroupName(fileRaw);
            Sheet sheet = reqWorkbook.getSheetAt(0);
            excelUtil.execute(sheet, mgroupName);
            reqWorkbook.close();

            moveFilesToDirectory(fileRaw);
        }catch (Exception e) {
            throw new GenericErrorException("Unable to process file", e);
        }
    }

    private void moveFilesToDirectory(File fileRaw){
        try {
            Files.createDirectories(Paths.get(System.getProperty("user.home") + "/JTGM MGroup/Processed"));
            File newFile = new File(System.getProperty("user.home") + "/JTGM MGroup/Processed/" + fileRaw.getName());
            File oldFile = new File(fileRaw.getPath());

            log.info("[INFO] Moving files from Raw folder to processed.");
            Files.move(oldFile.toPath(), newFile.toPath(), REPLACE_EXISTING);
        }catch (FileAlreadyExistsException ex){
            log.error("[ERROR] File already exist, will not transfer.");
        }catch (IOException ex){
            throw new GenericErrorException("Unable to process file", ex);
        }
    }

    private String getMgroupName(File file) {
        String name  = file.getName().replaceAll("[\\[\\](){}]",";");
        return name.split(";")[0];
    }
}
