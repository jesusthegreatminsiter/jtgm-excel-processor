package org.jtgm.conf;

import org.jtgm.core.DefaultExcelExtractor;
import org.jtgm.core.ExcelExtractor;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class CoreConfig {

    @Bean
    ExcelExtractor excelExtractor(HeaderProperties headerProperties){
        return new DefaultExcelExtractor(headerProperties);
    }
}
