package com.adcapricornio.excel_reports.common.exceptions;

public class ExcelReportException extends RuntimeException {

    public ExcelReportException(String message) {
        super(message);
    }

    public ExcelReportException(String message, Throwable cause) {
        super(message, cause);
    }

}
