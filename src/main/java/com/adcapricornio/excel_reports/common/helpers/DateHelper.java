package com.adcapricornio.excel_reports.common.helpers;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;

public class DateHelper {

    private static final DateTimeFormatter DATE_TIME_WITH_SECONDS =
            DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss");

    private static final DateTimeFormatter DATE_TIME_WITHOUT_SECONDS =
            DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm");

    private DateHelper() {
    }

    public static String currentDateTime() {
        ZoneId peruZoneId = ZoneId.of("America/Lima");
        LocalDateTime currentDateTime = LocalDateTime.now(peruZoneId);
        DateTimeFormatter format = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss");
        return currentDateTime.format(format);
    }

    public static DateTimeFormatter getDateTimeFormatter(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }

        String normalizedValue = value.trim();
        return normalizedValue.matches("\\d{2}/\\d{2}/\\d{4} \\d{2}:\\d{2}:\\d{2}")
                ? DATE_TIME_WITH_SECONDS
                : DATE_TIME_WITHOUT_SECONDS;

    }

}
