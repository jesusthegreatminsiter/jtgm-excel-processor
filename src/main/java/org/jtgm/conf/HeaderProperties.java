package org.jtgm.conf;

import lombok.Getter;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.stereotype.Component;

@Getter
@Component
@EnableConfigurationProperties
@ConfigurationProperties(prefix = "excel.header")
public class HeaderProperties {
    private String mgroup;
    private String name;
    private String leader;
    private String date;
}
