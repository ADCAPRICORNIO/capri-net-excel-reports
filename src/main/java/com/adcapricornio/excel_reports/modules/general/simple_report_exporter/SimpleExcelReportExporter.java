package com.adcapricornio.excel_reports.modules.general.simple_report_exporter;

import com.adcapricornio.excel_reports.common.constants.DataTypes;
import com.adcapricornio.excel_reports.common.exceptions.ExcelReportException;
import com.adcapricornio.excel_reports.common.helpers.DateHelper;
import com.adcapricornio.excel_reports.modules.general.simple_report_exporter.helpers.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;

public class SimpleExcelReportExporter {

    public void build(SimpleExcelReportInput data, OutputStream outputStream) {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {

            int currentRowNumber = 0;
            Sheet sheet = workbook.createSheet("Informe");

            /*====================================================*/
            /*                     FILTERS SECTION                */
            /*====================================================*/

            /*===================================*/
            /*          FILTERS HEADER           */
            /*===================================*/

            Row sheetFiltersHeader = sheet.createRow(0);

            Cell cellFiltersHeaderI = sheetFiltersHeader.createCell(0);
            cellFiltersHeaderI.setCellValue("FILTROS DE BUSQUEDA");
            cellFiltersHeaderI.setCellStyle(FiltersHeaderStyle.getInitialCellStyle(workbook));
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

            for (int i = 1; i < 4; i++) {
                Cell cellFiltersHeaderM = sheetFiltersHeader.createCell(i);
                cellFiltersHeaderM.setCellStyle(FiltersHeaderStyle.getMiddleCellStyle(workbook));
            }

            Cell cellFiltersHeaderF = sheetFiltersHeader.createCell(4);
            cellFiltersHeaderF.setCellStyle(FiltersHeaderStyle.getFinalCellStyle(workbook));

            currentRowNumber += 1;

            /*===================================*/
            /*          FILTERS CONTENT          */
            /*===================================*/

            var filters = new ArrayList<>(data.getFilters());
            filters.add(new SimpleExcelReportInput.Filter("FECHA CREACIÓN", DateHelper.currentDateTime()));

            for (int i = 0; i < filters.size(); i++) {

                Row sheetFiltersContent = sheet.createRow(i + 1);

                Cell keyCellFilterContent = sheetFiltersContent.createCell(0);
                keyCellFilterContent.setCellValue(filters.get(i).getKey() + ":");
                keyCellFilterContent.setCellStyle(FiltersContentStyle.getKeyCellStyle(workbook));

                Cell cellFiltersContentI = sheetFiltersContent.createCell(1);
                cellFiltersContentI.setCellValue(filters.get(i).getValue());
                cellFiltersContentI.setCellStyle(FiltersContentStyle.getInitialCellStyle(workbook));
                sheet.addMergedRegion(new CellRangeAddress(i + 1, i + 1, 1, 4));

                for (int x = 2; x < 4; x++) {
                    Cell cellFiltersContentM = sheetFiltersContent.createCell(x);
                    cellFiltersContentM.setCellStyle(FiltersContentStyle.getMiddleCellStyle(workbook));
                }

                Cell cellFiltersContentF = sheetFiltersContent.createCell(4);
                cellFiltersContentF.setCellStyle(FiltersContentStyle.getFinalCellStyle(workbook));

            }

            currentRowNumber += filters.size();

            /*====================================================*/
            /*                     DATA SECTION                   */
            /*====================================================*/

            /*===================================*/
            /*                 TITLE             */
            /*===================================*/

            currentRowNumber += 1;

            Row sheetTitle = sheet.createRow(currentRowNumber);

            Cell cellTitleI = sheetTitle.createCell(0);
            cellTitleI.setCellValue(data.getTitle());
            cellTitleI.setCellStyle(TitleStyle.getInitialCellStyle(workbook));
            sheet.addMergedRegion(new CellRangeAddress(currentRowNumber, currentRowNumber, 0, data.getData().getHeaders().size()));

            for (int x = 1; x < data.getData().getHeaders().size(); x++) {
                Cell cellTitleM = sheetTitle.createCell(x);
                cellTitleM.setCellStyle(TitleStyle.getMiddleCellStyle(workbook));
            }

            Cell cellTitleF = sheetTitle.createCell(data.getData().getHeaders().size());
            cellTitleF.setCellStyle(TitleStyle.getFinalCellStyle(workbook));

            currentRowNumber += 1;

            /*===================================*/
            /*             DATA HEADER           */
            /*===================================*/

            Row sheetDataHeader = sheet.createRow(currentRowNumber);

            Cell cellIndexHeader = sheetDataHeader.createCell(0);
            cellIndexHeader.setCellValue("ITEM");
            cellIndexHeader.setCellStyle(DataHeaderStyle.getCellStyle(workbook));

            for (int i = 0; i < data.getData().getHeaders().size(); i++) {
                sheet.setColumnWidth(i + 1, data.getData().getHeaders().get(i).getWidth() * 100);
                Cell cellHeader = sheetDataHeader.createCell(i + 1);
                cellHeader.setCellValue(data.getData().getHeaders().get(i).getLabel());
                cellHeader.setCellStyle(DataHeaderStyle.getCellStyle(workbook));
            }

            /*===================================*/
            /*                 DATA              */
            /*===================================*/

            DataStyleCache dataStyle = new DataStyleCache(workbook);

            for (int i = 0; i < data.getData().getRows().size(); i++) {
                var currentRow = data.getData().getRows().get(i);

                currentRowNumber += 1;
                Row dataSection = sheet.createRow(currentRowNumber);

                Cell cellIndexData = dataSection.createCell(0);
                cellIndexData.setCellValue(i + 1);
                cellIndexData.setCellStyle(DataStyle.getIndexCellStyle(workbook));

                for (int x = 0; x < currentRow.getItems().size(); x++) {
                    Cell cellData = dataSection.createCell(x + 1);

                    String textAlign = data.getData().getHeaders().get(x).getTextAlign();
                    String dataType = data.getData().getHeaders().get(x).getDataType();
                    String backgroundColor = currentRow.getMetadata().getBackgroundColor();
                    if (backgroundColor == null || backgroundColor.isEmpty()) backgroundColor = "#ffffff";

                    var currentValue = currentRow.getItems().get(x).getValue();
                    if (currentValue == null) currentValue = "";

                    if (currentValue.equals("")) {
                        cellData.setCellValue("");
                    } else {
                        switch (dataType) {
                            case DataTypes.INTEGER -> cellData.setCellValue(Integer.parseInt(currentValue.toString()));
                            case DataTypes.DOUBLE -> cellData.setCellValue(Double.parseDouble(currentValue.toString()));
                            case DataTypes.DATE -> {
                                SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
                                Date dateValue = dateFormat.parse(currentValue.toString());
                                cellData.setCellValue(dateValue);
                            }
                            case DataTypes.DATETIME -> {
                                DateTimeFormatter formatter = DateHelper.getDateTimeFormatter(currentValue.toString());
                                LocalDateTime dateTimeValue =
                                        LocalDateTime.parse(currentValue.toString(), formatter);
                                cellData.setCellValue(currentValue.toString());
                            }
                            case DataTypes.STRING -> cellData.setCellValue(currentValue.toString());
                        }
                    }

                    cellData.setCellStyle(dataStyle.getStyle(dataType, textAlign, true, true, backgroundColor));

                }
            }

            /*====================================================*/
            /*                     EXPORT DATA                    */
            /*====================================================*/

            workbook.write(outputStream);

        } catch (IOException | ParseException e) {
            throw new ExcelReportException("Error generating simple report", e);
        }
    }

}