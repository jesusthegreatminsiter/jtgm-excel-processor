package org.jtgm.presentation;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.jtgm.core.service.ExcelExtractor;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import java.io.File;

@RestController
@RequestMapping("/jtgm/excel")
@Slf4j
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelExtractor excelExtractor;

    @GetMapping(path = "/mgroup",  produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<String> saveMgroup() {
        try{
            log.info("[START] Initiating process of MGroup files to staging excel.");
            File folder = new File(System.getProperty("user.home") + "/JTGM MGroup/Raw");
            File[] listOfFiles = folder.listFiles();

            for (File file : listOfFiles) {
                if (file.isFile()) {
                    log.info("[INFO] File Name {}", file.getName());
                    excelExtractor.extract(file);
                }
            }

            log.info("[END] Done processing the MGroup files to staging excel.");
            return new ResponseEntity<>("Extraction finished", HttpStatus.OK);
        }catch (Exception ex){
            return new ResponseEntity<>("Failed to extract the excel file", HttpStatus.INTERNAL_SERVER_ERROR);
        }

    }
}
