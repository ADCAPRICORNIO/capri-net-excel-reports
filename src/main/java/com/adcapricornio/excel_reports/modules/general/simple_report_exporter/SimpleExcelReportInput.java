package com.adcapricornio.excel_reports.modules.general.simple_report_exporter;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class SimpleExcelReportInput {

    private String filename;
    private String title;
    private List<Filter> filters;
    private Data data;

    @Getter
    @Setter
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Filter {
        private String key;
        private String value;
    }

    @Getter
    @Setter
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Data {
        private List<Header> headers;
        private List<Row> rows;
    }

    @Getter
    @Setter
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Row {
        private List<Item> items;
        private Metadata metadata;
    }

    @Getter
    @Setter
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Metadata {
        private String backgroundColor;
    }

    @Getter
    @Setter
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Item {
        private String key;
        private Object value;
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
