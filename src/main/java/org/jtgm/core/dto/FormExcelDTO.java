package org.jtgm.core.dto;

import lombok.*;
import org.apache.poi.ss.usermodel.Cell;
import org.jtgm.conf.HeaderProperties;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Builder
@Getter
@Setter
@AllArgsConstructor
public class FormExcelDTO {
    private List<String> attendees;
    private String mgroupLeader;
    private Date date;
    private List<String> others;

    public static FormExcelDTO buildFormExcel(CellFinderDTO cellFinder, HeaderProperties headerProperties) {
        return builder()
                .attendees(getList(cellFinder, headerProperties.getAttendance()))
                .mgroupLeader(getStringValue(cellFinder, headerProperties.getLeader()))
                .date(getDateValue(cellFinder, headerProperties.getDate()))
                .others(getList(cellFinder, headerProperties.getOtherAttendee()))
                .build();
    }

    private static List<String> getList(CellFinderDTO cellFinder, String headerProperties) {
        String cell = getStringValue(cellFinder, headerProperties);
        List<String> finArr = new ArrayList<>();
        if(cell != null) {
            String[] tempArr = cell.split(";");
            for (String s : tempArr) {
                if (s != null && !s.isEmpty() && s != "null" && s != "") {
                    finArr.add(s);
                }
            }
        }
        return finArr;
    }

    private static String getStringValue(CellFinderDTO cellFinder, String headerName){
        try{
            int cellIndex = cellFinder.getFoundHeaderMap().get(headerName);
            Cell currentCell = cellFinder.getRow().getCell(cellIndex);

            return String.valueOf(currentCell);
        }catch (NullPointerException ex){
            return null;
        }
    }

    private static Date getDateValue(CellFinderDTO cellFinder, String headerName){
        try{
            int cellIndex = cellFinder.getFoundHeaderMap().get(headerName);
            Cell currentCell = cellFinder.getRow().getCell(cellIndex);
            return currentCell.getDateCellValue();
        }catch (Exception ex){
            return null;
        }
    }
}
