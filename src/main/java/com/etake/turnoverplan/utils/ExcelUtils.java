package com.etake.turnoverplan.utils;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;

import static com.etake.turnoverplan.utils.CalculationUtils.DEFAULT_ROUNDING_MODE;
import static com.etake.turnoverplan.utils.CalculationUtils.DEFAULT_SCALE;
import static java.util.Objects.nonNull;

@NoArgsConstructor(access = AccessLevel.NONE)
public final class ExcelUtils {

    public static <T> void setValue(final Row row, final int col, final T value) {
        final Cell cell = row.getCell(col);
        if (nonNull(value)) {
            if (value instanceof Number number) {
                final BigDecimal bigDecimal = BigDecimal.valueOf(number.doubleValue());
                if (bigDecimal.compareTo(BigDecimal.ZERO) > 0) {
                    cell.setCellValue(bigDecimal.doubleValue());
                }
            } else {
                cell.setCellValue(value.toString());
            }
        }
    }

    public static void createCells(final Sheet sheet, final int endRow, final int endCol) {
        for (int i = 0; i <= endRow; i++) {
            final Row row = sheet.createRow(i);
            for (int j = 0; j <= endCol; j++) {
                row.createCell(j);
            }
        }
    }

    public static void applyCellStyle(final Sheet sheet,
                                      final CellStyle cellStyle,
                                      final int startRow,
                                      final int startColumn,
                                      final int endRow,
                                      final int endColumn) {
        for (int i = startRow; i <= endRow; i++) {
            final Row row = sheet.getRow(i);
            for (int j = startColumn; j <= endColumn; j++) {
                final Cell cell = row.getCell(j);
                cell.setCellStyle(cellStyle);
            }
        }
    }

    public static void applyCellStyle(final Sheet sheet,
                                      final CellStyle cellStyle,
                                      final int row,
                                      final int col) {
        applyCellStyle(sheet, cellStyle, row, col, row, col);
    }

    public static void applyDataFormat(final Workbook workbook,
                                       final Sheet sheet,
                                       final int startRow,
                                       final int startColumn,
                                       final int endRow,
                                       final int endColumn,
                                       final String dataFormatStyle) {
        final DataFormat dataFormat = workbook.createDataFormat();
        short format = dataFormat.getFormat(dataFormatStyle);
        for (int i = startRow; i <= endRow; i++) {
            final Row row = sheet.getRow(i);
            for (int j = startColumn; j <= endColumn; j++) {
                final Cell cell = row.getCell(j);
                final CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.cloneStyleFrom(cell.getCellStyle());
                cellStyle.setDataFormat(format);
                ExcelUtils.applyCellStyle(sheet, cellStyle, i, j);
            }
        }
    }

    public static void applyDataFormat(final Workbook workbook,
                                       final Sheet sheet,
                                       final int row,
                                       final int col,
                                       final String dataFormatStyle) {
        applyDataFormat(workbook, sheet, row, col, row, col, dataFormatStyle);
    }

    public static void applyBackgroundColor(final Workbook workbook,
                                            final Sheet sheet,
                                            final int startRow,
                                            final int startColumn,
                                            final int endRow,
                                            final int endColumn,
                                            final IndexedColors color) {
        for (int i = startRow; i <= endRow; i++) {
            final Row row = sheet.getRow(i);
            for (int j = startColumn; j <= endColumn; j++) {
                final Cell cell = row.getCell(j);
                final CellStyle cellStyle = workbook.createCellStyle();
                if (nonNull(cell)) {
                    cellStyle.cloneStyleFrom(cell.getCellStyle());
                } else {
                    row.createCell(j);
                }
                cellStyle.setFillForegroundColor(color.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                ExcelUtils.applyCellStyle(sheet, cellStyle, i, j);
            }
        }
    }

    public static void applyBorders(final Workbook workbook,
                                    final Sheet sheet,
                                    final int startRow,
                                    final int startColumn,
                                    final int endRow,
                                    final int endColumn,
                                    final BorderStyle borderStyle) {
        for (int i = startRow; i <= endRow; i++) {
            final Row row = sheet.getRow(i);
            for (int j = startColumn; j <= endColumn; j++) {
                final Cell cell = row.getCell(j);
                final CellStyle cellStyle = workbook.createCellStyle();
                if (nonNull(cell)) {
                    cellStyle.cloneStyleFrom(cell.getCellStyle());
                } else {
                    row.createCell(j);
                }
                cellStyle.setBorderRight(borderStyle);
                cellStyle.setBorderLeft(borderStyle);
                cellStyle.setBorderTop(borderStyle);
                cellStyle.setBorderBottom(borderStyle);
                ExcelUtils.applyCellStyle(sheet, cellStyle, i, j);
            }
        }
    }

    public static void autosizeColumns(final Sheet sheet, final FormulaEvaluator evaluator) {
        final int startRow = sheet.getTopRow();
        final int startColumn = sheet.getRow(startRow).getFirstCellNum();
        final int endRow = sheet.getLastRowNum() - 1;
        final int endColumn = sheet.getRow(endRow).getLastCellNum() - 1;

        for (int i = startColumn; i <= endColumn; i++) {
            int maxLength = 0;
            for (int j = startRow + 1; j <= endRow; j++) {
                final Cell cell = sheet.getRow(j).getCell(i);
                if (cell.isPartOfArrayFormulaGroup()) {
                    continue;
                }
                final int currentLength = getCellValueLength(evaluator, cell);
                if (currentLength > maxLength) {
                    maxLength = currentLength;
                }
            }
            sheet.setColumnWidth(i, (maxLength + 2) * 255);
        }
    }

    private static int getCellValueLength(final FormulaEvaluator evaluator, final Cell cell) {
        return  switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().length();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue()).length();
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue()).length();
            case FORMULA -> getFormulaResultNumberLength(evaluator, cell);
            case BLANK -> 0;
            default -> throw new IllegalStateException("Cannot get length from cell value");
        };
    }

    private static int getFormulaResultNumberLength(final FormulaEvaluator evaluator, final Cell cell) {
        final BigDecimal bigDecimal = BigDecimal.valueOf(evaluator.evaluate(cell).getNumberValue())
                .setScale(DEFAULT_SCALE, DEFAULT_ROUNDING_MODE);
        return String.valueOf(bigDecimal.doubleValue()).length();
    }
}
