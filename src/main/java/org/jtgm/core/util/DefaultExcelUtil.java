package org.jtgm.core.util;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jtgm.conf.HeaderProperties;
import org.jtgm.core.dto.CellFinderDTO;
import org.jtgm.core.dto.FormExcelDTO;

import java.io.*;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.*;

import static org.jtgm.core.dto.CellFinderDTO.buildCellFinder;
import static org.jtgm.core.dto.FormExcelDTO.buildFormExcel;

@RequiredArgsConstructor
@Slf4j
public class DefaultExcelUtil implements ExcelUtil {

    private final HeaderProperties headerProperties;
    private static final int HEADER_ROW_NUMBER = 0;
    private static final String RESOURCE_NAME = "/file/Transactional.xlsx";

    @Override
    public void execute(Sheet sheet, String mgroupName) {
        try {
            HashMap<String, Integer> headers = getHeaders(sheet);
            List<FormExcelDTO> formExcelList = getInfoFromExcel(sheet, headers);
            String nameDate = new SimpleDateFormat("MM-dd-yyyy").format(getFridayOfWeek(new Date())) + " Staging.xlsx";
            File outputFile = new File(System.getProperty("user.home") + "/" + nameDate);

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
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void processRows(String mgroupName,
                             Workbook resWorkbook,
                             Sheet sheet,
                             FormExcelDTO formExcelDTO,
                             List<String> toProcess,
                             Boolean isOther) {
        if(!toProcess.isEmpty()) {
            List<String> attendee = formExcelDTO.getAttendees();
            for(int i = 0; i<attendee.size(); i++ ){
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

                String[] attendeeDet = attendee.get(i).split("-");

                Cell cell3 = row.createCell(3);
                cell3.setCellValue(Long.valueOf(removeSpaces(attendeeDet[0])));

                Cell cell4 = row.createCell(4);
                cell4.setCellValue(removeSpaces(attendeeDet[1]));

                if(isOther){
                    Cell cell5 = row.createCell(5);
                    cell5.setCellValue("Yes");
                }

                Cell cell6 = row.createCell(6);
                cell6.setCellValue(computeWeekNumber(formExcelDTO.getDate()));

                Cell cell7 = row.createCell(7);
                cell7.setCellValue(getFridayOfWeek(formExcelDTO.getDate()));
                cell7.setCellStyle(cellStyle);

                Cell cell8 = row.createCell(8);
                cell8.setCellValue(LocalDate.now());
                cell8.setCellStyle(cellStyle);
            }
        }
    }

    private int computeWeekNumber(Date date) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        return calendar.get(Calendar.WEEK_OF_YEAR);
    }

    private Date getFridayOfWeek(Date date) {
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.WEEK_OF_YEAR, computeWeekNumber(date));
        calendar.set(Calendar.DAY_OF_WEEK, Calendar.FRIDAY);

        return calendar.getTime();
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

    private String removeSpaces(String toFormat){
        return toFormat.replaceAll("\\s", "");
    }

    public static void copyFile(File out) throws Exception {
        Path path = FileSystems.getDefault().getPath(out.getPath());
        Files.copy(DefaultExcelUtil.class.getResourceAsStream(RESOURCE_NAME), path, StandardCopyOption.REPLACE_EXISTING);
    }
}
