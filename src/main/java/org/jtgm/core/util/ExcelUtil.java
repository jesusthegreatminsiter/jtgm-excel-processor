package org.jtgm.core.util;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jtgm.conf.HeaderProperties;
import org.jtgm.core.dto.CellFinderDTO;
import org.jtgm.core.dto.FormExcelDTO;
import org.jtgm.core.exception.GenericErrorException;

import java.io.*;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.*;

import static org.jtgm.core.dto.CellFinderDTO.buildCellFinder;
import static org.jtgm.core.dto.FormExcelDTO.buildFormExcel;
import static org.jtgm.core.util.GenericUtil.*;

@RequiredArgsConstructor
@Slf4j
public class ExcelUtil {

    private final ValidationUtil validationUtil;
    private final HeaderProperties headerProperties;
    private static final int HEADER_ROW_NUMBER = 0;

    public void execute(Sheet sheet, String mgroupName) {
        try {
            HashMap<String, Integer> headers = getHeaders(sheet);
            List<FormExcelDTO> formExcelList = getInfoFromExcel(sheet, headers);

            File outputFile = generateOutputFile(new Date());
            FileInputStream file = new FileInputStream(outputFile);
            Workbook resWorkbook = new XSSFWorkbook(file);
            Sheet sheetRes = resWorkbook.getSheetAt(0);

            for(int j = 0; j < formExcelList.size(); j++) {
                FormExcelDTO formExcelDTO = formExcelList.get(j);
                String path = outputFile.getPath();
                processRows(mgroupName, resWorkbook, sheetRes, formExcelDTO, formExcelDTO.getAttendees(), false, path);
                processRows(mgroupName, resWorkbook, sheetRes, formExcelDTO, formExcelDTO.getOthers(), true, path);
            }
            FileOutputStream fos = new FileOutputStream(outputFile);
            resWorkbook.write(fos);
            fos.close();
            resWorkbook.close();
        } catch (Exception e) {
            e.printStackTrace();
            throw new GenericErrorException("[ERROR] Failed to continue processing file", e);
        }
    }

    private static File generateOutputFile(Date date) {
        try{
            String weekOfYear = new SimpleDateFormat("MM-dd-yyyy").format(getFridayOfWeek(date));
            String nameDate = System.getProperty("user.home") + "/JTGM MGroup/" + weekOfYear + " Staging.xlsx";
            File outputFile = new File(nameDate);

            if (!outputFile.exists()) {
                copyFile(outputFile);
            }

            return outputFile;
        }catch(Exception e ){
            throw new GenericErrorException("[ERROR] Failed to generate output file", e);
        }
    }

    private void processRows(String mgroupName,
                             Workbook resWorkbook,
                             Sheet sheet,
                             FormExcelDTO formExcelDTO,
                             List<String> toProcess,
                             Boolean isOther,
                             String fileName) {
        try {
            int thisWeekNum = computeWeekNumber(new Date());
            if(!toProcess.isEmpty()) {
                for(int i = 0; i<= toProcess.size() - 1; i++ ){
                    String[] attendeeDet = toProcess.get(i).split(" - ");
                    int weekNumber = computeWeekNumber(formExcelDTO.getDate());

                    Row row;
                    CellStyle cellStyle;
                    File outputFile = null;
                    Workbook extraWorkbook = null;
                    CreationHelper createHelper;

                    if (thisWeekNum != weekNumber){
                        outputFile = generateOutputFile(formExcelDTO.getDate());
                        FileInputStream file = new FileInputStream(outputFile);
                        extraWorkbook = new XSSFWorkbook(file);
                        Sheet sheetRes = extraWorkbook.getSheetAt(0);

                        fileName = outputFile.getPath();
                        row = sheetRes.createRow(sheetRes.getLastRowNum() + 1);
                        cellStyle = extraWorkbook.createCellStyle();
                        createHelper = extraWorkbook.getCreationHelper();
                    }else{
                        row = sheet.createRow(sheet.getLastRowNum() + 1);
                        cellStyle = resWorkbook.createCellStyle();
                        createHelper = resWorkbook.getCreationHelper();
                    }
                    boolean doesExist =  validationUtil.validate(attendeeDet, weekNumber, isOther, mgroupName, fileName);
                    if(doesExist){
                        continue;
                    }

                    Cell cell0 = row.createCell(0);


                    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm-dd-yyyy"));

                    cell0.setCellValue(formExcelDTO.getDate());
                    cell0.setCellStyle(cellStyle);

                    Cell cell1 = row.createCell(1);
                    cell1.setCellValue(mgroupName);

                    Cell cell2 = row.createCell(2);
                    cell2.setCellValue(formExcelDTO.getMgroupLeader());

                    Cell cell4 = row.createCell(4);

                    if(isOther){
                        Cell cell5 = row.createCell(5);
                        cell5.setCellValue("Yes");

                        cell4.setCellValue(attendeeDet[0]);

                    }else{
                        Cell cell3 = row.createCell(3);
                        cell3.setCellValue(Long.parseLong(attendeeDet[0]));

                        cell4.setCellValue(attendeeDet[1]);
                    }

                    Cell cell6 = row.createCell(6);
                    cell6.setCellValue(weekNumber);

                    Cell cell7 = row.createCell(7);
                    cell7.setCellValue(getFridayOfWeek(formExcelDTO.getDate()));
                    cell7.setCellStyle(cellStyle);

                    Cell cell8 = row.createCell(8);
                    cell8.setCellValue(LocalDate.now());
                    cell8.setCellStyle(cellStyle);

                    if(outputFile!=null){
                        FileOutputStream fos = new FileOutputStream(outputFile);
                        extraWorkbook.write(fos);
                        fos.close();
                        extraWorkbook.close();
                    }
                }
            }
        } catch (Exception e) {
            throw new GenericErrorException("[ERROR] Failed to process file", e);
        }
    }


    private List<FormExcelDTO>  getInfoFromExcel(Sheet sheet, HashMap<String, Integer> headers) {
        List<FormExcelDTO> formExcelDTOList = new ArrayList<>();
        for(Row row : sheet) {
            if(row.getRowNum() != HEADER_ROW_NUMBER){
                CellFinderDTO cellFinder = buildCellFinder(headers, row);
                formExcelDTOList.add(buildFormExcel(cellFinder, headerProperties));
            }
        }
        return formExcelDTOList;
    }

    private HashMap<String, Integer> getHeaders(Sheet sheet) {
        HashMap<String, Integer> headerMap = new HashMap<>();

        Row row = sheet.getRow(0);
        for(Cell cell : row) {
            headerMap.put(
                    removeSpaces(cell.getStringCellValue().toLowerCase()),
                    cell.getColumnIndex()
            );
        }
        return headerMap;
    }
}
