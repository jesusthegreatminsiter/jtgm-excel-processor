package org.jtgm.conf;

import org.jtgm.core.service.impl.DefaultExcelExtractor;
import org.jtgm.core.service.ExcelExtractor;
import org.jtgm.core.util.ExcelUtil;
import org.jtgm.core.util.ValidationUtil;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class CoreConfig {

    @Bean
    ExcelExtractor excelExtractor(ExcelUtil excelUtil){
        return new DefaultExcelExtractor(excelUtil);
    }

    @Bean
    ExcelUtil excelUtil(ValidationUtil validationUtil, HeaderProperties headerProperties){
        return new ExcelUtil(validationUtil, headerProperties);
    }

    @Bean
    ValidationUtil validationUtil(){
        return new ValidationUtil();
    }
}
