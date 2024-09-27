package com.etake.turnoverplan.utils;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.stream.IntStream;

@NoArgsConstructor(access = AccessLevel.NONE)
public final class ExcelFormulaUtils {

    public static String getDynFormula(final String numeratorRange, final String denominatorRange) {
        final String numeratorPart = "SUMIFS(%s,$A:$A,INDIRECT(\"$B\"&ROW()),$C:$C,INDIRECT(\"$C\"&ROW()))".formatted(numeratorRange);
        final String denominatorPart = "SUMIFS(%s,$A:$A,INDIRECT(\"$B\"&ROW()),$C:$C,INDIRECT(\"$C\"&ROW()))".formatted(denominatorRange);
        return "IF(OR(%s = 0, %s = 0), 0, %s / %s - 1)".formatted(numeratorPart, denominatorPart, numeratorPart, denominatorPart);
    }

    public static String getAvgPlanFormula(final String avgColumnName, final String dynColumnName) {
        return "INDIRECT(\"%s\"&ROW()) * (1 + INDIRECT(\"%s\"&ROW()))".formatted(avgColumnName, dynColumnName);
    }

    public static String getPlanFormula(final String avgPlanColumnName, final String daysRange) {
        final String daysPart = "SUMIFS(%s,$A:$A,INDIRECT(\"$B\"&ROW()),$C:$C,INDIRECT(\"$C\"&ROW()))".formatted(daysRange);
        return "INDIRECT(\"%s\"&ROW()) * %s".formatted(avgPlanColumnName, daysPart);
    }

    public static String getCategoriesPlanFormula(final String sumRange) {
        return "SUMIF(data!$C:$C,INDIRECT(\"A\"&ROW()),data!%s)".formatted(sumRange);
    }

    public static String getTotalSumFormula(final int startRow, final int endRow) {
        return "SUBTOTAL(9, INDIRECT(SUBSTITUTE(ADDRESS(ROW(), COLUMN(), 4), ROW(), \"\") & \"%d:\" & SUBSTITUTE(ADDRESS(ROW(), COLUMN(), 4), ROW(), \"\") & \"%d\"))".formatted(startRow, endRow);
    }

    public static void setFormula(final String formula, final Sheet sheet, final Integer startRow,
                                  final Integer endRow, final Integer startColumn, final Integer endColumn) {
        for (int i = startRow; i <= endRow; i++) {
            for (int j = startColumn; j <= endColumn; j++) {
                sheet.getRow(i).getCell(j).setCellFormula(formula);
            }
        }
    }

    public static void setFormula(final String formula, final Sheet sheet, final Integer startRow,
                                  final Integer endRow, final Integer column) {
        IntStream.range(startRow, endRow)
                .forEach(i -> sheet.getRow(i).getCell(column).setCellFormula(formula));
    }
}
