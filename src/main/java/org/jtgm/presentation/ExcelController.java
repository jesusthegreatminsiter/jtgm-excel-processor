package org.jtgm.presentation;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.jtgm.core.ExcelExtractor;
import org.jtgm.presentation.exception.RestControllerException;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/jtgm/excel")
@Slf4j
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelExtractor excelExtractor;

    @PostMapping(path = "/hello",  produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<String> sampleEp(@RequestParam("name") String name) {
        try{
            log.info("[START] Hitting a test endpoint");
            return new ResponseEntity<>("Hello, " + name, HttpStatus.OK);
        }catch(Exception e){
            log.error("[ERROR] Error in hitting the test endpoint ");
            throw new RestControllerException(e.getMessage());
        }
    }
}
