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
            String year = String.valueOf(Calendar.getInstance().get(Calendar.YEAR));
            String nameDate = System.getProperty("user.home") + "/JTGM MGroup/" + year  + " Report.xlsx";
            File outputFile = new File(nameDate);

            if (!outputFile.exists()) {
                copyFile(outputFile);
            }

            FileInputStream file = new FileInputStream(outputFile);
            Workbook resWorkbook = new XSSFWorkbook(file);
            Sheet sheetRes = resWorkbook.getSheetAt(0);

            for(int j = 0; j < formExcelList.size(); j++) {
                FormExcelDTO formExcelDTO = formExcelList.get(j);
                processRows(mgroupName, resWorkbook, sheetRes, formExcelDTO, formExcelDTO.getAttendees(), false);
                processRows(mgroupName, resWorkbook, sheetRes, formExcelDTO, formExcelDTO.getOthers(), true);
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

    private void processRows(String mgroupName,
                             Workbook resWorkbook,
                             Sheet sheet,
                             FormExcelDTO formExcelDTO,
                             List<String> toProcess,
                             Boolean isOther) {

        if(!toProcess.isEmpty()) {
            for(int i = 0; i <= toProcess.size() - 1; i++ ){
                String[] attendeeDet = toProcess.get(i).split(" - ");
                int weekNumber = computeWeekNumber(formExcelDTO.getDate());
                boolean doesExist =  validationUtil.validate(attendeeDet, weekNumber, isOther, mgroupName);

                if(doesExist){
                    continue;
                }

                Row row = sheet.createRow(sheet.getLastRowNum() + 1);

                Cell cell0 = row.createCell(0);
                CellStyle cellStyle = resWorkbook.createCellStyle();
                CreationHelper createHelper = resWorkbook.getCreationHelper();
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

                    String[] newAttendee = attendeeDet[0].split("\\r?\\n");
                    newAttendee = Arrays.stream(newAttendee).filter(Objects::nonNull)
                            .filter(s -> !Objects.equals(s, ""))
                            .toArray(String[]::new);
                    cell4.setCellValue(newAttendee[0]);

                    if(newAttendee.length > 1) {
                        for (int y = 1; y <= newAttendee.length - 1; y++) {
                            Row row_ = sheet.createRow(sheet.getLastRowNum() + 1);

                            Cell cell01 = row_.createCell(0);
                            cell01.setCellValue(formExcelDTO.getDate());
                            cell01.setCellStyle(cellStyle);

                            Cell cell1_ = row_.createCell(1);
                            cell1_.setCellValue(mgroupName);

                            Cell cell2_ = row_.createCell(2);
                            cell2_.setCellValue(formExcelDTO.getMgroupLeader());

                            Cell cell4_ = row_.createCell(4);
                            cell4_.setCellValue(newAttendee[y]);

                            Cell cell6_ = row_.createCell(6);
                            cell6_.setCellValue(weekNumber);

                            Cell cell7_ = row_.createCell(7);
                            cell7_.setCellValue(getFridayOfWeek(formExcelDTO.getDate()));
                            cell7_.setCellStyle(cellStyle);

                            Cell cell8_ = row_.createCell(8);
                            cell8_.setCellValue(LocalDate.now());
                            cell8_.setCellStyle(cellStyle);
                            Cell cell5_ = row_.createCell(5);
                            cell5_.setCellValue("Yes");
                        }

                    }
                }else{
                    Cell cell3 = row.createCell(3);
                    cell3.setCellValue(Long.valueOf(attendeeDet[0]));

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
            }
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
