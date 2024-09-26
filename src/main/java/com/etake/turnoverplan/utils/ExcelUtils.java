package com.etake.turnoverplan.utils;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import static java.util.Objects.nonNull;

@NoArgsConstructor(access = AccessLevel.NONE)
public final class ExcelUtils {

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
}
