package com.adcapricornio.excel_reports.modules.general.multi_sheet_report_exporter.helpers;

import org.apache.poi.ss.usermodel.*;

public class TitleStyle {

    public static CellStyle getInitialCellStyle(Workbook workbook) {
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
