package com.adcapricornio.excel_reports.modules.customized.general_series.helpers;

import org.apache.poi.ss.usermodel.*;

public class GeneralDataStyle {

    public static CellStyle getTitleCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET TEXT ALIGN
        style.setAlignment(HorizontalAlignment.CENTER);

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

    public static CellStyle getKeyCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

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

    public static CellStyle getContentLeftCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET BORDER
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    public static CellStyle getContentRightCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET TEXT ALIGN
        style.setAlignment(HorizontalAlignment.RIGHT);

        // SET BORDER
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    public static CellStyle getContentCenterCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // SET TEXT ALIGN
        style.setAlignment(HorizontalAlignment.CENTER);

        // SET BORDER
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

}
