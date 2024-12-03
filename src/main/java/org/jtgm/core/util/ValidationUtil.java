package org.jtgm.core.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jtgm.core.dto.RowDTO;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static org.jtgm.core.util.GenericUtil.getFridayOfWeek;

@Slf4j
public class ValidationUtil {

    public boolean validate(String[] personId, int weekNumber, boolean isOthers, String mgroup)  {
        boolean doesExist = false;
        try{
            log.info("[INFO] Excel validation is starting...");
            RowDTO rowDTO = RowDTO.builder()
                    .fullName(isOthers ? personId[0] : personId[1])
                    .weekNumber(weekNumber)
                    .mgroup(mgroup)
                    .build();

            List<RowDTO> listFromExcel = getListFromExcel();
            int size = listFromExcel.size();
            if (size != 0) {
                for(int x = 0; x <= size - 1; x++){
                    RowDTO rowData = listFromExcel.get(x);
                    if(rowData.getFullName().equalsIgnoreCase(rowDTO.getFullName())) {
                        if(rowData.getWeekNumber() == rowDTO.getWeekNumber()){
                            if(rowData.getMgroup().equalsIgnoreCase(rowDTO.getMgroup())) {
                                doesExist = true;
                                break;
                            }
                        }
                    }
                }
            }
            return doesExist;
        }catch (Exception ex){
            ex.printStackTrace();
            throw new RuntimeException("[ERROR] Cannot validate the excel if exist");
        }
    }

    public List<RowDTO> getListFromExcel() throws Exception {
        log.info("[INFO] Fetching data from existing excel...");
        String weekOfYear = new SimpleDateFormat("MM-dd-yyyy").format(getFridayOfWeek(new Date()));
        String nameDate = System.getProperty("user.home") + "/" + weekOfYear + " Staging.xlsx";
        File outputFile = new File(nameDate);

        if (!outputFile.exists()) {
            return  null;
        }

        FileInputStream file = new FileInputStream(outputFile);
        Workbook resWorkbook = new XSSFWorkbook(file);
        Sheet sheetRes = resWorkbook.getSheetAt(0);
        int rowTotal = sheetRes.getLastRowNum();

        List<RowDTO> rowDTOList = new ArrayList<>();

      for(int i = 0; i <= rowTotal; i++){
          Row row = sheetRes.getRow(i);
            if(row.getRowNum() != 0){
                rowDTOList.add(RowDTO.builder()
                        .fullName(row.getCell(4).getStringCellValue())
                        .weekNumber((int) row.getCell(6).getNumericCellValue())
                        .mgroup(row.getCell(1).getStringCellValue())
                        .build());
            }
        }
        return rowDTOList;
    }
}
