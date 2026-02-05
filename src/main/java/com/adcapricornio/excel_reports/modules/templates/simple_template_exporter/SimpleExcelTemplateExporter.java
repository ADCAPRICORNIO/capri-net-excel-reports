package com.adcapricornio.excel_reports.modules.templates.simple_template_exporter;

import com.adcapricornio.excel_reports.common.exceptions.ExcelReportException;
import com.adcapricornio.excel_reports.modules.templates.simple_template_exporter.helpers.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;

public class SimpleExcelTemplateExporter {

    public void build(SimpleExcelTemplateInput data, OutputStream outputStream) {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {

            int currentRowNumber = 0;
            Sheet sheet = workbook.createSheet("Informe");

            /*====================================================*/
            /*                     DATA SECTION                   */
            /*====================================================*/

            /*===================================*/
            /*             DATA HEADER           */
            /*===================================*/

            Row sheetDataHeader = sheet.createRow(currentRowNumber);

            for (int i = 0; i < data.getData().getHeaders().size(); i++) {
                sheet.setColumnWidth(i, data.getData().getHeaders().get(i).getWidth() * 100);
                Cell cellHeader = sheetDataHeader.createCell(i);
                cellHeader.setCellValue(data.getData().getHeaders().get(i).getLabel());
                cellHeader.setCellStyle(DataHeaderStyle.getCellStyle(workbook));
            }

            /*====================================================*/
            /*                     EXPORT DATA                    */
            /*====================================================*/

            workbook.write(outputStream);

        } catch (IOException e) {
            throw new ExcelReportException("Error generating simple report", e);
        }

    }

}
