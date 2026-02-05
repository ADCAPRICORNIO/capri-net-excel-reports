package com.adcapricornio.excel_reports.common.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class CellStyles {

    private Workbook workbook;

    private String dataType;
    private String textAlign;
    private Boolean wrapText = false;
    private Boolean shrinkToFit = false;
    private Boolean borderLeft = false;
    private Boolean borderRight = false;
    private Boolean borderTop = false;
    private Boolean borderBottom = false;
    private String backgroundColorHex;

    public CellStyles(Workbook workbook) {
        this.workbook = workbook;
    }

    public CellStyles setDataType(String value) {
        this.dataType = value;
        return this;
    }

    public CellStyles setTextAlign(String value) {
        this.textAlign = value;
        return this;
    }

    public CellStyles setWrapText(Boolean value) {
        this.wrapText = value;
        return this;
    }

    public CellStyles setShrinkToFit(Boolean value) {
        this.shrinkToFit = value;
        return this;
    }

    public CellStyles setBorderTop(Boolean value) {
        this.borderTop = value;
        return this;
    }

    public CellStyles setBorderBottom(Boolean value) {
        this.borderBottom = value;
        return this;
    }

    public CellStyles setBorderLeft(Boolean value) {
        this.borderLeft = value;
        return this;
    }

    public CellStyles setBorderRight(Boolean value) {
        this.borderRight = value;
        return this;
    }

    public CellStyles setBackgroundColor(String hex) {
        this.backgroundColorHex = hex;
        return this;
    }


    public CellStyle build() {
        CellStyle style;

        if (workbook instanceof SXSSFWorkbook sxssf) {
            style = sxssf.getXSSFWorkbook().createCellStyle();
        } else {
            style = workbook.createCellStyle();
        }

        DataFormat format = workbook.createDataFormat();

        // TEXT ALIGN
        switch (this.textAlign) {
            case "L":
                style.setAlignment(HorizontalAlignment.LEFT);
                break;
            case "C":
                style.setAlignment(HorizontalAlignment.CENTER);
                break;
            case "R":
                style.setAlignment(HorizontalAlignment.RIGHT);
                break;
            default:
                style.setAlignment(HorizontalAlignment.LEFT);
        }

        switch (this.dataType) {
            case "DD":
                style.setDataFormat(format.getFormat("dd/MM/yyyy"));
                break;
            case "DT":
                style.setDataFormat(format.getFormat("dd/MM/yyyy HH:mm:ss"));
                break;
            case "I":
                style.setDataFormat(format.getFormat("0"));
                break;
            case "D":
                style.setDataFormat(format.getFormat("0.00"));
                break;
            default:
                style.setDataFormat(format.getFormat("@"));
        }

        style.setWrapText(this.wrapText);

        if (this.borderTop) style.setBorderTop(BorderStyle.THIN);
        if (this.borderBottom) style.setBorderBottom(BorderStyle.THIN);
        if (this.borderLeft) style.setBorderLeft(BorderStyle.THIN);
        if (this.borderRight) style.setBorderRight(BorderStyle.THIN);

        if (this.backgroundColorHex != null && !this.backgroundColorHex.isEmpty() && style instanceof org.apache.poi.xssf.usermodel.XSSFCellStyle xssfStyle) {
            String hex = this.backgroundColorHex.replace("#", "");
            byte[] rgb = new byte[]{
                    (byte) Integer.parseInt(hex.substring(0, 2), 16),
                    (byte) Integer.parseInt(hex.substring(2, 4), 16),
                    (byte) Integer.parseInt(hex.substring(4, 6), 16)
            };

            var color = new org.apache.poi.xssf.usermodel.XSSFColor(rgb, null);

            xssfStyle.setFillForegroundColor(color);
            xssfStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
        }

        return style;
    }

}
