package com.adcapricornio.excel_reports.modules.customized.general_series;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class GeneralSeriesReportInput {

    private GeneralData generalData;
    private CustomsValue customsValue;
    private MerchandiseSummary merchandiseSummary;
    private DutySettlement dutySettlement;
    private List<DutySettlementTracking> dutySettlementTracking;
    private List<ValueSummary> valueSummary;
    private String invoiceNumber;

    @Getter
    @Setter
    public static class GeneralData {
        private String year;
        private String client;
        private String netK;
        private String grossK;
        private String declaration;
        private String warehouse;
        private String orderNumber;
        private String declarationDate;
        private String reference;
        private String ruc;
        private String bulks;
        private String channel;
        private String customs;
        private String vapour;
        private String uf;
        private String withdrawalDate;
        private String regimen;
        private String eta;
        private String arrivalDate;
        private String uc;
        private String cancellationDate;
        private String series;
    }

    @Getter
    @Setter
    public static class CustomsValue {
        private String fobCurrency;
        private String fobValue;
        private String fobValue2;
        private String freightCurrency;
        private String freightValue;
        private String freightValue2;
        private String insuranceCurrency;
        private String insuranceValue;
        private String insuranceValue2;
        private String adjustmentCurrency;
        private String adjustmentValue;
        private String adjustmentValue2;
        private String cifCurrency;
        private String cifValue;
        private String cifValue2;
    }

    @Getter
    @Setter
    public static class MerchandiseSummary {
        private String usdExchangeRate;
        private String currencyExchangeRate;
    }

    @Getter
    @Setter
    public static class DutySettlement {
        private String adValoremValue1;
        private String adValoremValue2;
        private String adValoremValue3;
        private String adValoremValue4;
        private String adValoremValue5;
        private String adSpecificValue1;
        private String adSpecificValue2;
        private String adSpecificValue3;
        private String adSpecificValue4;
        private String adSpecificValue5;
        private String surchargeValue1;
        private String surchargeValue2;
        private String surchargeValue3;
        private String surchargeValue4;
        private String surchargeValue5;
        private String iscValue1;
        private String iscValue2;
        private String iscValue3;
        private String iscValue4;
        private String iscValue5;
        private String igvValue1;
        private String igvValue2;
        private String igvValue3;
        private String igvValue4;
        private String igvValue5;
        private String ipmValue1;
        private String ipmValue2;
        private String ipmValue3;
        private String ipmValue4;
        private String ipmValue5;
        private String antidumpingValue1;
        private String antidumpingValue2;
        private String antidumpingValue3;
        private String antidumpingValue4;
        private String antidumpingValue5;
        private String compoundInterestValue1;
        private String compoundInterestValue2;
        private String compoundInterestValue3;
        private String compoundInterestValue4;
        private String compoundInterestValue5;
        private String customsClearanceFeeValue1;
        private String customsClearanceFeeValue2;
        private String customsClearanceFeeValue3;
        private String customsClearanceFeeValue4;
        private String customsClearanceFeeValue5;
        private String totalValue1;
        private String totalValue2;
        private String totalValue3;
        private String totalValue4;
        private String totalValue5;
        private String perceptionValue1;
        private String perceptionValue2;
        private String perceptionValue3;
        private String totalPaymentAmount;
    }

    @Getter
    @Setter
    public static class DutySettlementTracking {
        private String series;
        private String nandina;
        private String fobUsd;
        private String freightUsd;
        private String insuranceUsd;
        private String expenseUsd;
        private String cifUsd;
        private String adValorem;
        private String adValoremPercentage;
        private String specificDescription;
        private String isc;
        private String iscPercentage;
        private String igv;
        private String igvPercentage;
        private String ipm;
        private String ipmPercentage;
        private String antidumping;
        private String ds35;
        private String ds35Percentage;
        private String totalTaxUsd;
        private String perception;
        private String perceptionPercentage;
        private String originPortCode;
        private String originPort;
        private String naladisa;
        private String nabandina;
        private String tpi;
        private String tpn;
        private String releasingCode;
        private String grossK;
        private String netK;
        private String type;
        private String commercialUnit;
        private String type2;
        private String bulks;
        private String description01;
        private String description02;
        private String description03;
        private String description04;
        private String description05;
        private String invoiceNumber;
    }

    @Getter
    @Setter
    public static class ValueSummary {
        private String percentage;
        private String series;
        private String fob;
        private String freight;
        private String insurance;
        private String cif;
        private String adValorem;
        private String igv;
        private String ipm;
        private String others;
    }

}
