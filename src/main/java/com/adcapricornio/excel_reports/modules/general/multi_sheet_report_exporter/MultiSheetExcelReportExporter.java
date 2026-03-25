package com.adcapricornio.excel_reports.modules.general.multi_sheet_report_exporter;

import com.adcapricornio.excel_reports.common.constants.DataTypes;
import com.adcapricornio.excel_reports.common.exceptions.ExcelReportException;
import com.adcapricornio.excel_reports.common.helpers.DateHelper;
import com.adcapricornio.excel_reports.modules.general.multi_sheet_report_exporter.helpers.FiltersContentStyle;
import com.adcapricornio.excel_reports.modules.general.simple_report_exporter.helpers.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;

public class MultiSheetExcelReportExporter {

    public void build(MultiSheetExcelReportInput data, OutputStream outputStream) {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {

            for (MultiSheetExcelReportInput.Worksheet worksheet : data.getWorksheets()) {

                int currentRowNumber = 0;
                Sheet sheet = workbook.createSheet(worksheet.getName());

                /*====================================================*/
                /*                     FILTERS SECTION                */
                /*====================================================*/

                /*===================================*/
                /*          FILTERS HEADER           */
                /*===================================*/

                if (!worksheet.getShowFiltersTitle().equals(0)) {
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
                }

                /*===================================*/
                /*          FILTERS CONTENT          */
                /*===================================*/

                var filterColumns = new ArrayList<>(worksheet.getFilterColumns());

                if (!filterColumns.isEmpty()) {

                    var firstColumnFilters = new ArrayList<>(filterColumns.getFirst().getFilters());
                    firstColumnFilters.add(new MultiSheetExcelReportInput.Filter("FECHA CREACIÓN", DateHelper.currentDateTime()));

                    filterColumns.set(0, new MultiSheetExcelReportInput.FilterColumn(firstColumnFilters));

                    var filterColumnMaxSize = worksheet.getFilterColumns().stream()
                            .mapToInt(fc -> fc.getFilters().size())
                            .max()
                            .orElse(0);

                    for (int i = 0; i < filterColumnMaxSize; i++) {
                        Row sheetFiltersContent = sheet.createRow(i + 1);
                        int columnIndex = 0;

                        for (var filterColumn : filterColumns) {
                            var filters = filterColumn.getFilters();

                            if (i < filters.size()) {
                                Cell keyCellFilterContent = sheetFiltersContent.createCell(columnIndex);
                                keyCellFilterContent.setCellValue(filters.get(i).getKey() + ":");
                                keyCellFilterContent.setCellStyle(FiltersContentStyle.getKeyCellStyle(workbook));

                                Cell cellFiltersContentI = sheetFiltersContent.createCell(1 + columnIndex);
                                if (filters.get(i).getValue() != null && !filters.get(i).getValue().equals("null"))
                                    cellFiltersContentI.setCellValue(filters.get(i).getValue());
                                cellFiltersContentI.setCellStyle(FiltersContentStyle.getInitialCellStyle(workbook));
                                sheet.addMergedRegion(new CellRangeAddress(i + 1, i + 1, 1 + columnIndex, 2 + columnIndex));

                                Cell cellFiltersContentF = sheetFiltersContent.createCell(2 + columnIndex);
                                cellFiltersContentF.setCellStyle(FiltersContentStyle.getFinalCellStyle(workbook));
                            }

                            columnIndex += 3;
                        }
                    }

                    currentRowNumber += filterColumnMaxSize + 1;

                }

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
                sheet.addMergedRegion(new CellRangeAddress(currentRowNumber, currentRowNumber, 0, worksheet.getData().getHeaders().size()));

                for (int x = 1; x < worksheet.getData().getHeaders().size(); x++) {
                    Cell cellTitleM = sheetTitle.createCell(x);
                    cellTitleM.setCellStyle(TitleStyle.getMiddleCellStyle(workbook));
                }

                Cell cellTitleF = sheetTitle.createCell(worksheet.getData().getHeaders().size());
                cellTitleF.setCellStyle(TitleStyle.getFinalCellStyle(workbook));

                currentRowNumber += 1;

                /*===================================*/
                /*             DATA HEADER           */
                /*===================================*/

                Row sheetDataHeader = sheet.createRow(currentRowNumber);

                Cell cellIndexHeader = sheetDataHeader.createCell(0);
                cellIndexHeader.setCellValue("ITEM");
                cellIndexHeader.setCellStyle(DataHeaderStyle.getCellStyle(workbook));

                for (int i = 0; i < worksheet.getData().getHeaders().size(); i++) {
                    sheet.setColumnWidth(i + 1, worksheet.getData().getHeaders().get(i).getWidth() * 100);
                    Cell cellHeader = sheetDataHeader.createCell(i + 1);
                    cellHeader.setCellValue(worksheet.getData().getHeaders().get(i).getLabel());
                    cellHeader.setCellStyle(DataHeaderStyle.getCellStyle(workbook));
                }

                /*===================================*/
                /*                 DATA              */
                /*===================================*/

                DataStyleCache dataStyle = new DataStyleCache(workbook);

                for (int i = 0; i < worksheet.getData().getRows().size(); i++) {
                    var currentRow = worksheet.getData().getRows().get(i);

                    currentRowNumber += 1;
                    Row dataSection = sheet.createRow(currentRowNumber);

                    Cell cellIndexData = dataSection.createCell(0);
                    cellIndexData.setCellValue(i + 1);
                    cellIndexData.setCellStyle(DataStyle.getIndexCellStyle(workbook));

                    for (int x = 0; x < currentRow.getItems().size(); x++) {
                        Cell cellData = dataSection.createCell(x + 1);

                        String textAlign = worksheet.getData().getHeaders().get(x).getTextAlign();
                        String dataType = worksheet.getData().getHeaders().get(x).getDataType();
                        String backgroundColor = currentRow.getMetadata().getBackgroundColor();
                        if (backgroundColor == null || backgroundColor.isEmpty()) backgroundColor = "#ffffff";

                        var currentValue = currentRow.getItems().get(x).getValue();
                        if (currentValue == null) currentValue = "";

                        if (currentValue.equals("")) {
                            cellData.setCellValue("");
                        } else {
                            switch (dataType) {
                                case DataTypes.INTEGER ->
                                        cellData.setCellValue(Integer.parseInt(currentValue.toString()));
                                case DataTypes.DOUBLE ->
                                        cellData.setCellValue(Double.parseDouble(currentValue.toString()));
                                case DataTypes.DATE -> {
                                    SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
                                    Date dateValue = dateFormat.parse(currentValue.toString());
                                    cellData.setCellValue(dateValue);
                                }
                                case DataTypes.DATETIME -> {
                                    DateTimeFormatter formatter =
                                            DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss");
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

            }

            workbook.write(outputStream);

        } catch (Exception e) {
            throw new ExcelReportException("Error generating multi sheet report", e);
        }
    }

}
