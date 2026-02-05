package com.adcapricornio.excel_reports.modules.templates.simple_template_exporter.helpers;

import org.apache.poi.ss.usermodel.*;

public class DataHeaderStyle {

    public static CellStyle getCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET TEXT ALIGN
        style.setAlignment(HorizontalAlignment.CENTER);

        // SET FONT WEIGHT
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        // SET BACKGROUND
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // SET BORDER
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }
}
