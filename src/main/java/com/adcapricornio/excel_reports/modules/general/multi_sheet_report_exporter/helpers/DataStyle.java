package com.adcapricornio.excel_reports.modules.general.multi_sheet_report_exporter.helpers;

import org.apache.poi.ss.usermodel.*;

public class DataStyle {

    public static CellStyle getIndexCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET TEXT ALIGN
        style.setAlignment(HorizontalAlignment.CENTER);

        // SET FONT WEIGHT
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        // SET BORDER
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    public static CellStyle getDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET BORDER
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

}
