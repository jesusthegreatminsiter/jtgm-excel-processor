package org.jtgm.core;

import com.monitorjbl.xlsx.StreamingReader;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.jtgm.conf.HeaderProperties;
import org.jtgm.core.dto.CellFinderDTO;
import org.jtgm.core.dto.FormExcelDTO;
import org.jtgm.core.exception.GenericErrorException;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import static org.jtgm.core.dto.CellFinderDTO.buildCellFinder;
import static org.jtgm.core.dto.FormExcelDTO.buildFormExcel;

@RequiredArgsConstructor
public class DefaultExcelExtractor implements ExcelExtractor{

    private final HeaderProperties headerProperties;
    private static final int HEADER_ROW_NUMBER = 0;

    @Override
    public void extract(MultipartFile file) {
        try {
            Workbook workbook = StreamingReader.builder().open(file.getInputStream());

            Sheet sheet = workbook.getSheetAt(0);

            HashMap<String, Integer> headers = getHeaders(sheet);

            List<FormExcelDTO> formExcelDTOList = new ArrayList<>();
            for(Row row : sheet){
                if(row.getRowNum() != HEADER_ROW_NUMBER){
                    CellFinderDTO cellFinder = buildCellFinder(headers, row);
                    formExcelDTOList.add(buildFormExcel(cellFinder, headerProperties));
                }
            }


            File outputFile = new File(System.getProperty("user.dir") +"/final.xlsx");

            FileOutputStream fos = new FileOutputStream(outputFile);
            workbook.write(fos);

        }catch (Exception e) {
            e.printStackTrace();
            throw new GenericErrorException("Error occur", e);
        }
    }

    private HashMap<String, Integer> getHeaders(Sheet sheet){
        HashMap<String, Integer> headerMap = new HashMap<>();

        Row row = sheet.rowIterator().next();
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
}
