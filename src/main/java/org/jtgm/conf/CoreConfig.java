package org.jtgm.conf;

import org.jtgm.core.service.DefaultExcelExtractor;
import org.jtgm.core.service.ExcelExtractor;
import org.jtgm.core.util.DefaultExcelUtil;
import org.jtgm.core.util.ExcelUtil;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class CoreConfig {

    @Bean
    ExcelExtractor excelExtractor(ExcelUtil excelUtil){
        return new DefaultExcelExtractor(excelUtil);
    }

    @Bean
    ExcelUtil excelUtil(HeaderProperties headerProperties){
        return new DefaultExcelUtil(headerProperties);
    }
}
