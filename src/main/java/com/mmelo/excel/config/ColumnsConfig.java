package com.mmelo.excel.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

import java.util.List;

@Configuration
@Data
@ConfigurationProperties("table")
public class ColumnsConfig {
    private Integer start;
    private List<String> columns;
}