package org.jtgm.core.dto;

import lombok.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.jtgm.conf.HeaderProperties;

@Builder
@Getter
@Setter
@AllArgsConstructor
public class FormExcelDTO {
    private String name;
    private String mgroupLeader;
    private String date;
    private String mgroupName;

    public static FormExcelDTO buildFormExcel(CellFinderDTO cellFinder, HeaderProperties headerProperties) {
        return builder()
                .name(getCellValue(cellFinder, headerProperties.getName()))
                .mgroupLeader(getCellValue(cellFinder, headerProperties.getLeader()))
                .mgroupName(getCellValue(cellFinder, headerProperties.getMgroup()))
                .date(getCellValue(cellFinder, headerProperties.getDate()))
                .build();
    }

    private static String getCellValue(CellFinderDTO cellFinder, String headerName){
        try{
            int cellIndex = cellFinder.getFoundHeaderMap().get(headerName);
            Cell currentCell = cellFinder.getRow().getCell(cellIndex);

            String cellValue = null;

            switch(currentCell.getCellType()){
                case STRING:
                    cellValue = currentCell.getStringCellValue();
                    break;
                case NUMERIC:
                    if(DateUtil.isCellDateFormatted(currentCell)){
                        cellValue = String.valueOf(currentCell.getDateCellValue());
                    }else{
                        cellValue = String.valueOf(currentCell.getNumericCellValue());
                    }
                    break;
            }
            return cellValue;
        }catch (NullPointerException ex){
            return null;
        }
    }

}
