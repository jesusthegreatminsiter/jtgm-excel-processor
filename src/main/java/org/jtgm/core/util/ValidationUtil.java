package org.jtgm.core.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jtgm.core.dto.RowDTO;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;


@Slf4j
public class ValidationUtil {

    public boolean validate(String[] personId, int weekNumber, boolean isOthers, String mgroup)  {
        boolean doesExist = false;
        try{
            String name = isOthers && personId.length > 1 ? personId[0] : personId[1];
            log.info(String.valueOf(personId.length));
            RowDTO rowDTO = RowDTO.builder()
                    .fullName(name)
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
        String year = String.valueOf(Calendar.getInstance().get(Calendar.YEAR));
        String nameDate = System.getProperty("user.home") + "/JTGM MGroup/" + year  + " Report.xlsx";
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
