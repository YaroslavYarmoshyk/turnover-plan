package com.etake.turnoverplan.service;

import com.etake.turnoverplan.config.properties.SystemConfigurationProperties;
import com.etake.turnoverplan.model.Quarter;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Component;

import java.time.YearMonth;

@Component
@RequiredArgsConstructor
public class PeriodProvider {
    private final SystemConfigurationProperties systemConfigurationProperties;

    public static YearMonth getYearMonth(final YearMonth yearMonth,
                                         final Integer yearOffset,
                                         final Integer monthOffset) {
        return yearMonth.plusYears(yearOffset).plusMonths(monthOffset);
    }

    public Quarter getPlannedQuarter() {
        return getQuarter(0);
    }

    public Quarter getPrevQuarter() {
        return getQuarter(-1);
    }

    private Quarter getQuarter(final Integer yearOffset) {
        int i = 0;
        final String actualMonthName = getOffsetMontYearName(yearOffset, i++);
        final String firstMonthName = getOffsetMontYearName(yearOffset, i++);
        final String secondMonthName = getOffsetMontYearName(yearOffset, i++);
        final String thirdMonthName = getOffsetMontYearName(yearOffset, i);
        return new Quarter(actualMonthName, firstMonthName, secondMonthName, thirdMonthName);
    }

    private String getOffsetMontYearName(final Integer yearOffset, final Integer monthOffset) {
        final YearMonth currentYearMonth = YearMonth.of(systemConfigurationProperties.year(), systemConfigurationProperties.month());
        final YearMonth yearMonth = currentYearMonth.plusYears(yearOffset).plusMonths(monthOffset);

        return String.format("%s - %s", getMonthPart(yearMonth.getMonthValue()), yearMonth.getYear());
    }

    private static String getMonthPart(final Integer month) {
        return switch (month) {
            case 1 -> "Січ";
            case 2 -> "Лют";
            case 3 -> "Бер";
            case 4 -> "Кві";
            case 5 -> "Тра";
            case 6 -> "Чер";
            case 7 -> "Лип";
            case 8 -> "Сер";
            case 9 -> "Вер";
            case 10 -> "Жов";
            case 11 -> "Лис";
            case 12 -> "Гру";
            default -> throw new IllegalStateException("Cannot resolve month name");
        };
    }
}
