package com.adcapricornio.excel_reports.modules.templates.simple_template_exporter;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class SimpleExcelTemplateInput {

    private String filename;
    private Data data;

    @Getter
    @Setter
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Data {
        private List<Header> headers;
    }

    @Getter
    @Setter
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Header {
        private String key;
        private String label;

        private Integer width;

        private String dataType;
        private String textAlign;
    }
}
