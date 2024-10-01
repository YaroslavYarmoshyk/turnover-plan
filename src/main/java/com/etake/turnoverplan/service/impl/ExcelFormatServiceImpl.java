package com.etake.turnoverplan.service.impl;

import com.etake.turnoverplan.config.indices.ResultSheetColumnIndices;
import com.etake.turnoverplan.config.properties.ColumnIndices;
import com.etake.turnoverplan.service.ExcelFormatService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.stereotype.Service;

import java.util.HashSet;
import java.util.Set;

import static com.etake.turnoverplan.utils.Constants.FIRST_ROW_INDEX;
import static com.etake.turnoverplan.utils.Constants.INITIAL_VALUE_ROW_INDEX;
import static com.etake.turnoverplan.utils.Constants.SECOND_ROW_INDEX;
import static com.etake.turnoverplan.utils.ExcelUtils.applyCellStyle;
import static com.etake.turnoverplan.utils.ExcelUtils.autosizeColumns;

@Service
@RequiredArgsConstructor
public class ExcelFormatServiceImpl implements ExcelFormatService {
    private static final XSSFColor GREY_COLOR = new XSSFColor(new byte[]{(byte) 38, (byte) 38, (byte) 38}, new DefaultIndexedColorMap());
    private static final XSSFColor LIGHT_GREY_COLOR = new XSSFColor(new byte[]{(byte) 64, (byte) 64, (byte) 64}, new DefaultIndexedColorMap());
    private final ColumnIndices columnIndices;

    @Override
    public void formatDataSheet(final Workbook workbook, final Sheet sheet) {
        formatHeaders(workbook, sheet, columnIndices.dataSheet());

        final CellStyle defaultCellStyle = getDefaultCellStyle(workbook);
        final int endRowIndex = sheet.getLastRowNum() - 1;
        final int startColumnIndex = 0;
        final int endColumnIndex = sheet.getRow(FIRST_ROW_INDEX).getLastCellNum() - 1;
        applyCellStyle(sheet, defaultCellStyle, INITIAL_VALUE_ROW_INDEX, startColumnIndex, endRowIndex, endColumnIndex);

        final DataFormat dataFormat = workbook.createDataFormat();
        final CellStyle numberCellStyle = getDefaultCellStyle(workbook);
        numberCellStyle.setDataFormat(dataFormat.getFormat("_-* # ##0_-;-* # ##0_-;_-* \"-\"??_-;_-@_-"));
        final CellStyle percentageCellStyle = getDefaultCellStyle(workbook);
        percentageCellStyle.setDataFormat(dataFormat.getFormat("0.0%"));

        final Set<Integer> numberFormatIndices = columnIndices.dataSheet().getNumberFormatIndices();
        numberFormatIndices.forEach(i -> applyCellStyle(sheet, numberCellStyle, INITIAL_VALUE_ROW_INDEX, i, endRowIndex, i));
        columnIndices.dataSheet().getPercentageFormatIndices().forEach(i -> applyCellStyle(sheet, percentageCellStyle, INITIAL_VALUE_ROW_INDEX, i, endRowIndex, i));

        final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        autosizeColumns(sheet, evaluator);
    }

    private static void formatHeaders(final Workbook workbook, final Sheet sheet, final ResultSheetColumnIndices resultSheetColumnIndices) {
        final Row firstRow = sheet.getRow(FIRST_ROW_INDEX);
        final short defaultHeight = firstRow.getHeight();
        firstRow.setHeight((short) (defaultHeight * 2));
        final Row secondRow = sheet.getRow(SECOND_ROW_INDEX);
        secondRow.setHeight(defaultHeight);
        final CellStyle greyHeaderCellStyle = getHeaderCellStyle(workbook, GREY_COLOR);
        firstRow.cellIterator().forEachRemaining(cell -> cell.setCellStyle(greyHeaderCellStyle));

        final CellStyle lightGreyHeaderCellStyle = getHeaderCellStyle(workbook, LIGHT_GREY_COLOR);
        final Set<Integer> saleIndicatorIndices = new HashSet<>(resultSheetColumnIndices.getTurnoverColumnIndices());
        saleIndicatorIndices.addAll(resultSheetColumnIndices.getMarginColumnIndices());
        saleIndicatorIndices.forEach(i -> secondRow.getCell(i).setCellStyle(lightGreyHeaderCellStyle));
    }

    @Override
    public void formatCategoriesSheet(final Workbook workbook, final Sheet sheet) {

    }

    @Override
    public void formatStoresSheet(final Workbook workbook, final Sheet sheet) {

    }


    private static CellStyle getHeaderCellStyle(final Workbook workbook,
                                                final Color backgroundColor) {
        final CellStyle cellStyle = getDefaultHeaderCellStyle(workbook);
        cellStyle.setFillForegroundColor(backgroundColor);
        cellStyle.setFont(getFontByColor(workbook, IndexedColors.WHITE));
        return cellStyle;
    }

    private static CellStyle getDefaultHeaderCellStyle(final Workbook workbook) {
        final CellStyle cellStyle = getDefaultCellStyle(workbook);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setWrapText(true);
        return cellStyle;
    }

    private static CellStyle getDefaultCellStyle(final Workbook workbook) {
        final CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }


    private static Font getFontByColor(final Workbook workbook, final IndexedColors color) {
        final Font font = workbook.createFont();
        font.setColor(color.getIndex());
        font.setBold(true);
        return font;
    }
}
