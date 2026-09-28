package com.adcapricornio.excel_reports.modules.customized.general_series;

import com.adcapricornio.excel_reports.common.exceptions.ExcelReportException;
import com.adcapricornio.excel_reports.modules.customized.general_series.helpers.GeneralDataStyle;
import lombok.Builder;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class GeneralSeriesReportExporter {

    public void build(GeneralSeriesReportInput data, OutputStream outputStream) {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Informe");

            List<Integer> widthList = List.of(4000, 5000, 4000, 5000, 4000, 5000, 4000, 5000, 4000, 5000, 3000, 2000, 3000, 2000, 3000, 2000, 3000, 2000, 4000, 2000, 5000, 4000, 2000, 4000, 3000, 3000, 4000, 3000, 3000, 3000, 3000, 3000, 3000, 2000, 5000, 2000, 3000, 5000, 5000, 5000, 5000, 5000, 5000);

            for (int x = 0; x < widthList.size(); x++) {
                sheet.setColumnWidth(x, widthList.get(x));
            }

            CellStyle titleStyle = GeneralDataStyle.getTitleCellStyle(workbook);
            CellStyle keyStyle = GeneralDataStyle.getKeyCellStyle(workbook);
            CellStyle contentLeftStyle = GeneralDataStyle.getContentLeftCellStyle(workbook);
            CellStyle contentRightStyle = GeneralDataStyle.getContentRightCellStyle(workbook);
            CellStyle contentCenterStyle = GeneralDataStyle.getContentCenterCellStyle(workbook);

            /*=========================================================*/
            /*                     GENERAL DATA SECTION                */
            /*=========================================================*/

            /*==========================*/
            /*          TITLE           */
            /*==========================*/

            var generalData = data.getGeneralData();
            var customsValue = data.getCustomsValue();
            var merchandiseSummary = data.getMerchandiseSummary();
            var dutySettlement = data.getDutySettlement();

            if (generalData == null || customsValue == null || merchandiseSummary == null || dutySettlement == null)
                throw new ExcelReportException("Error al generar reporte Serie General");

            List<List<RowCellData>> cells = new ArrayList<>();

            cells.add(List.of(RowCellData.builder().index(0).value("DATOS GENERALES DEL PEDIDO").cellStyle(titleStyle).mergedRegion(new CellRange(0, 9)).build()));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("Año:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(generalData.getYear()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(2).value("Orden:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(3).value(generalData.getOrderNumber()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(4).value("Refer:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(5).value(generalData.getReference()).cellStyle(contentLeftStyle).mergedRegion(new CellRange(5, 7)).build(),
                    RowCellData.builder().index(8).value("Eta:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(9).value(generalData.getEta()).cellStyle(contentLeftStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("Cliente:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(generalData.getClient()).cellStyle(contentLeftStyle).mergedRegion(new CellRange(1, 3)).build(),
                    RowCellData.builder().index(4).value("Ruc:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(5).value(generalData.getRuc()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(6).value("Vapor:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(7).value(generalData.getVapour()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(8).value("F.Llegada:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(9).value(generalData.getArrivalDate()).cellStyle(contentLeftStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("K.Neto:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(generalData.getNetK()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(2).value("K.Bruto:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(3).value(generalData.getGrossK()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(4).value("Bultos:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(5).value(generalData.getBulks()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(6).value("U.F.:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(7).value(generalData.getUf()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value("U.C.").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(9).value(generalData.getUc()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("Declarac:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(generalData.getDeclaration()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(2).value("F.Declarac:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(3).value(generalData.getDeclarationDate()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(4).value("Canal:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(5).value(generalData.getChannel()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(6).value("Fch.Reti.:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(7).value(generalData.getWithdrawalDate()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(8).value("Fch.Canc.:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(9).value(generalData.getCancellationDate()).cellStyle(contentLeftStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("Almacen:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(generalData.getWarehouse()).cellStyle(contentLeftStyle).mergedRegion(new CellRange(1, 3)).build(),
                    RowCellData.builder().index(4).value("Aduana:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(5).value(generalData.getCustoms()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(6).value("Regimen:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(7).value(generalData.getRegimen()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value("Serie:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(9).value(generalData.getSeries()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of());
            cells.add(List.of(
                    RowCellData.builder().index(0).value("VALOR ADUANERO").cellStyle(titleStyle).mergedRegion(new CellRange(0, 3)).build(),
                    RowCellData.builder().index(5).value("LIQUIDACION DE DERECHOS").cellStyle(titleStyle).mergedRegion(new CellRange(5, 10)).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(1).value("Moneda").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(2).value("Divisas").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(3).value("Dolares:").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(5).value("").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(6).value("Liquidac.").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(7).value("Liberac.").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(8).value("Cant.Pagar").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(9).value("Liq.Aduanas").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(10).value("Soles").cellStyle(titleStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("Fob").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(customsValue.getFobCurrency()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(2).value(customsValue.getFobValue()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(3).value(customsValue.getFobValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(5).value("Ad.Valorem").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getAdValoremValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getAdValoremValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getAdValoremValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getAdValoremValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getAdValoremValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("Flete").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(customsValue.getFreightCurrency()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(2).value(customsValue.getFreightValue()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(3).value(customsValue.getFreightValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(5).value("Ad.Especifico").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getAdSpecificValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getAdSpecificValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getAdSpecificValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getAdSpecificValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getAdSpecificValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("Seguro").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(customsValue.getInsuranceCurrency()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(2).value(customsValue.getInsuranceValue()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(3).value(customsValue.getInsuranceValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(5).value("Sobretasa").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getSurchargeValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getSurchargeValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getSurchargeValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getSurchargeValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getSurchargeValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("Ajustes").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(customsValue.getAdjustmentCurrency()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(2).value(customsValue.getAdjustmentValue()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(3).value(customsValue.getAdjustmentValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(5).value("I.S.C").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getIscValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getIscValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getIscValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getIscValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getIscValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("CIF").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(customsValue.getCifCurrency()).cellStyle(contentLeftStyle).build(),
                    RowCellData.builder().index(2).value(customsValue.getCifValue()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(3).value(customsValue.getCifValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(5).value("I.G.V.").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getIgvValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getIgvValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getIgvValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getIgvValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getIgvValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(5).value("I.P.M.").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getIpmValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getIpmValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getIpmValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getIpmValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getIpmValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(1).value("RESUMEN DE LA MERCANCIA").cellStyle(titleStyle).mergedRegion(new CellRange(0, 3)).build(),
                    RowCellData.builder().index(5).value("Antidumping").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getAntidumpingValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getAntidumpingValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getAntidumpingValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getAntidumpingValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getAntidumpingValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("SINTOX").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value("").cellStyle(keyStyle).mergedRegion(new CellRange(1, 3)).build(),
                    RowCellData.builder().index(5).value("Interes Comp.").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getCompoundInterestValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getCompoundInterestValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getCompoundInterestValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getCompoundInterestValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getCompoundInterestValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value("").cellStyle(keyStyle).mergedRegion(new CellRange(1, 3)).build(),
                    RowCellData.builder().index(5).value("Tasa Desp.").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getCustomsClearanceFeeValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getCustomsClearanceFeeValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getCustomsClearanceFeeValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getCustomsClearanceFeeValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getCustomsClearanceFeeValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value("").cellStyle(keyStyle).mergedRegion(new CellRange(1, 3)).build(),
                    RowCellData.builder().index(5).value("Total").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value(dutySettlement.getTotalValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value(dutySettlement.getTotalValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getTotalValue3()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getTotalValue4()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getTotalValue5()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value("").cellStyle(keyStyle).mergedRegion(new CellRange(1, 3)).build(),
                    RowCellData.builder().index(5).value("Percepcion $").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value("").cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value("").cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getPerceptionValue1()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value(dutySettlement.getPerceptionValue2()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value(dutySettlement.getPerceptionValue3()).cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("TC.Dolar").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(merchandiseSummary.getUsdExchangeRate()).cellStyle(contentRightStyle).mergedRegion(new CellRange(1, 3)).build(),
                    RowCellData.builder().index(5).value("").cellStyle(keyStyle).mergedRegion(new CellRange(5, 10)).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("TC.Divisa").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(merchandiseSummary.getCurrencyExchangeRate()).cellStyle(contentRightStyle).mergedRegion(new CellRange(1, 3)).build(),
                    RowCellData.builder().index(5).value("Total").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(6).value("").cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(7).value("").cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(8).value(dutySettlement.getTotalPaymentAmount()).cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(9).value("").cellStyle(contentRightStyle).build(),
                    RowCellData.builder().index(10).value("").cellStyle(contentRightStyle).build()
            ));
            cells.add(List.of());
            cells.add(List.of(
                    RowCellData.builder().index(0).value("SEGUIMIENTO DE LIQUIDACION POR SERIE").cellStyle(titleStyle).mergedRegion(new CellRange(0, 10)).build()
            ));
            cells.add(List.of(
                    RowCellData.builder().index(0).value("SERIE").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(1).value("NANDINA").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(2).value("F.O.B. US$").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(3).value("FLETE US$").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(4).value("SEGURO US$").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(5).value("GASTOS US$").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(6).value("C.I.F. US$").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(7).value("AD-VALOREM").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(8).value("%").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(9).value("D.ESPECIF.").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(10).value("I.S.C").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(11).value("%").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(12).value("I.G.V").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(13).value("%").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(14).value("I.P.M").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(15).value("%").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(16).value("ANTIDUMP").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(17).value("DS 035 S/T").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(18).value("%").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(19).value("TOTAL TRIBUTOS US$").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(20).value("PERCEPCION S/").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(21).value("%").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(22).value("COD P.ORIGEN").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(23).value("P.ORIGEN").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(24).value("NALADISA").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(25).value("NABANDINA").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(26).value("T.P.I").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(27).value("T.P.N").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(28).value("C.LIB.").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(29).value("K.BRUTO").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(30).value("K.NETO").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(31).value("TIPO").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(32).value("U.COMER").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(33).value("TIPO").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(34).value("C.BULTOS").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(35).value("DESCRIPCION 1").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(36).value("DESCRIPCION 2").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(37).value("DESCRIPCION 3").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(38).value("DESCRIPCION 4").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(39).value("DESCRIPCION 5").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(40).value("NRO.FACTURA").cellStyle(titleStyle).build()
            ));

            data.getDutySettlementTracking().forEach(item -> cells.add(processDutySettlementTracking(item, contentLeftStyle, contentRightStyle, contentCenterStyle)));

            cells.add(List.of());
            cells.add(List.of(RowCellData.builder().index(0).value("RESUMEN DE VALORES").cellStyle(titleStyle).mergedRegion(new CellRange(0, 9)).build()));

            cells.add(List.of(
                    RowCellData.builder().index(0).value("PORCENTAJE").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(1).value("SERIE").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(2).value("FOB").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(3).value("FLETE").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(4).value("SEGURO").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(5).value("CIF").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(6).value("ADVALOREM").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(7).value("IGV").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(8).value("IPM").cellStyle(titleStyle).build(),
                    RowCellData.builder().index(9).value("OTROS").cellStyle(titleStyle).build()
            ));

            data.getValueSummary().forEach(item -> cells.add(processValueSummary(item, contentLeftStyle, contentRightStyle, contentCenterStyle)));

            cells.add(List.of());
            cells.add(List.of(
                    RowCellData.builder().index(0).value("FACTURA:").cellStyle(keyStyle).build(),
                    RowCellData.builder().index(1).value(data.getInvoiceNumber()).cellStyle(keyStyle).build()
            ));

            data.getValueSummary().forEach(item -> cells.add(processValueSummary(item, contentLeftStyle, contentRightStyle, contentCenterStyle)));


            this.multipleRowGenerator(sheet, cells);

            /*===========================*/
            /*          FIELDS           */
            /*===========================*/

            workbook.write(outputStream);

        } catch (IOException e) {
            throw new ExcelReportException("Error al generar el reporte de series generales", e);
        }
    }

    private List<RowCellData> processDutySettlementTracking(GeneralSeriesReportInput.DutySettlementTracking data, CellStyle leftStyle, CellStyle rightStyle, CellStyle centerStyle) {
        return List.of(
                RowCellData.builder().index(0).value(data.getSeries()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(1).value(data.getNandina()).cellStyle(centerStyle).build(),
                RowCellData.builder().index(2).value(data.getFobUsd()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(3).value(data.getFreightUsd()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(4).value(data.getInsuranceUsd()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(5).value(data.getExpenseUsd()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(6).value(data.getCifUsd()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(7).value(data.getAdValorem()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(8).value(data.getAdValoremPercentage()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(9).value(data.getSpecificDescription()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(10).value(data.getIsc()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(11).value(data.getIscPercentage()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(12).value(data.getIgv()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(13).value(data.getIgvPercentage()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(14).value(data.getIpm()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(15).value(data.getIpmPercentage()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(16).value(data.getAntidumping()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(17).value(data.getDs35()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(18).value(data.getDs35Percentage()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(19).value(data.getTotalTaxUsd()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(20).value(data.getPerception()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(21).value(data.getPerceptionPercentage()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(22).value(data.getOriginPortCode()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(23).value(data.getOriginPort()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(24).value(data.getNaladisa()).cellStyle(centerStyle).build(),
                RowCellData.builder().index(25).value(data.getNabandina()).cellStyle(centerStyle).build(),
                RowCellData.builder().index(26).value(data.getTpi()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(27).value(data.getTpn()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(28).value(data.getReleasingCode()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(29).value(data.getGrossK()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(30).value(data.getNetK()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(31).value(data.getType()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(32).value(data.getCommercialUnit()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(33).value(data.getType2()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(34).value(data.getBulks()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(35).value(data.getDescription01()).cellStyle(leftStyle).build(),
                RowCellData.builder().index(36).value(data.getDescription02()).cellStyle(leftStyle).build(),
                RowCellData.builder().index(37).value(data.getDescription03()).cellStyle(leftStyle).build(),
                RowCellData.builder().index(38).value(data.getDescription04()).cellStyle(leftStyle).build(),
                RowCellData.builder().index(39).value(data.getDescription05()).cellStyle(leftStyle).build(),
                RowCellData.builder().index(40).value(data.getInvoiceNumber()).cellStyle(leftStyle).build()
        );
    }

    private List<RowCellData> processValueSummary(GeneralSeriesReportInput.ValueSummary data, CellStyle leftStyle, CellStyle rightStyle, CellStyle centerStyle) {
        return List.of(
                RowCellData.builder().index(0).value(data.getPercentage()).cellStyle(centerStyle).build(),
                RowCellData.builder().index(1).value(data.getSeries()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(2).value(data.getFob()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(3).value(data.getFreight()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(4).value(data.getInsurance()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(5).value(data.getCif()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(6).value(data.getAdValorem()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(7).value(data.getIgv()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(8).value(data.getIpm()).cellStyle(rightStyle).build(),
                RowCellData.builder().index(9).value(data.getOthers()).cellStyle(rightStyle).build()
        );
    }

    private void multipleRowGenerator(Sheet sheet, List<List<RowCellData>> rows) {
        for (int i = 0; i < rows.size(); i++) {
            this.rowGenerator(sheet, i, rows.get(i));
        }
    }

    private void rowGenerator(Sheet sheet, Integer rowNumber, List<RowCellData> cells) {
        Row row = sheet.createRow(rowNumber);

        for (RowCellData cellData : cells) {
            Cell cell = row.createCell(cellData.getIndex());
            cell.setCellValue(cellData.getValue());
            cell.setCellStyle(cellData.getCellStyle());

            if (cellData.getMergedRegion() != null) {
                sheet.addMergedRegion(new CellRangeAddress(
                        rowNumber,
                        rowNumber,
                        cellData.getMergedRegion().getFirstColumn(),
                        cellData.getMergedRegion().getLastColumn()
                ));
            }
        }
    }

    @Getter
    @Builder
    private static class RowCellData {
        private Integer index;
        private String value;
        private CellStyle cellStyle;
        private CellRange mergedRegion;
    }

    @Getter
    private static class CellRange {
        private Integer firstColumn;
        private Integer lastColumn;

        public CellRange() {
        }

        public CellRange(Integer firstColumn, Integer lastColumn) {
            this.firstColumn = firstColumn;
            this.lastColumn = lastColumn;
        }
    }

}
