package com.adcapricornio.excel_reports.modules.templates.simple_template_exporter.helpers;

import org.apache.poi.ss.usermodel.*;

public class FiltersContentStyle {

    public static CellStyle getKeyCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET TEXT ALIGN
        style.setAlignment(HorizontalAlignment.LEFT);

        // SET FONT WEIGHT
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        // SET BORDER
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    public static CellStyle getInitialCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET TEXT ALIGN
        style.setAlignment(HorizontalAlignment.LEFT);

        // SET BORDER
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);

        return style;
    }

    public static CellStyle getMiddleCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET BORDER
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        return style;
    }

    public static CellStyle getFinalCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET BORDER
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

}
