package com.adcapricornio.excel_reports.modules.general.multi_sheet_report_exporter.helpers;

import com.adcapricornio.excel_reports.common.utils.CellStyles;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class DataStyleCache {

    private final ThreadLocal<Map<String, CellStyle>> threadLocalCache =
            ThreadLocal.withInitial(HashMap::new);
    private final Workbook workbook; // Workbook por request!

    public DataStyleCache(Workbook workbook) {
        this.workbook = workbook;
    }

    public CellStyle getStyle(
            String dataType,
            String textAlign,
            boolean borderLeft,
            boolean borderRight,
            String backgroundColorHex
    ) {

        String key = dataType + "|" + textAlign + "|" + borderLeft + "|" + borderRight;

        Map<String, CellStyle> cache = threadLocalCache.get();
        return cache.computeIfAbsent(key, k ->
                new CellStyles(workbook)
                        .setDataType(dataType)
                        .setTextAlign(textAlign)
                        .setWrapText(false)
                        .setBorderLeft(borderLeft)
                        .setBorderRight(borderRight)
                        .setBackgroundColor(backgroundColorHex)
                        .build()
        );
    }

    public void clear() {
        threadLocalCache.remove(); // Importante limpiar después del request
    }

}
