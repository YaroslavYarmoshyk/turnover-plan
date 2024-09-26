package com.etake.turnoverplan.service.impl;

import com.etake.turnoverplan.model.Quarter;
import com.etake.turnoverplan.model.StoreCategorySales;
import com.etake.turnoverplan.service.ExcelService;
import com.etake.turnoverplan.service.PeriodProvider;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.Collection;
import java.util.stream.IntStream;

import static com.etake.turnoverplan.utils.Constants.CATEGORY;
import static com.etake.turnoverplan.utils.Constants.SIMILAR_STORE;
import static com.etake.turnoverplan.utils.Constants.STORE;
import static com.etake.turnoverplan.utils.ExcelUtils.applyCellStyle;
import static java.util.Objects.isNull;
import static java.util.Objects.nonNull;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {
    private final PeriodProvider periodProvider;

    @Override
    public Workbook getWorkbook(final Collection<StoreCategorySales> sales) {
        final Workbook workbook = new SXSSFWorkbook();
        final Quarter prevQuarter = periodProvider.getPrevQuarter();
        final Quarter plannedQuarter = periodProvider.getPlannedQuarter();
        final Sheet dataSheet = getDataSheet(workbook, prevQuarter, plannedQuarter, sales);
        final Sheet storesPlan = workbook.createSheet("План по магазинам");
        final Sheet categoriesPlan = workbook.createSheet("План по категоріям");

        return workbook;
    }

    @Override
    public void writeWorkbook(final Workbook workbook, final String path) throws Exception {
        final FileOutputStream outputStream = new FileOutputStream(path);
        workbook.write(outputStream);
        workbook.close();
    }

    private Sheet getDataSheet(final Workbook workbook,
                               final Quarter prevQuarter,
                               final Quarter plannedQuarter,
                               final Collection<StoreCategorySales> sales) {
        final Sheet dataSheet = workbook.createSheet("data");
        createDataSheetHeader(workbook, prevQuarter, plannedQuarter, dataSheet);
        int startRow = 2;
        for (final StoreCategorySales storeCategorySales : sales) {
            final Row workRow = dataSheet.createRow(startRow++);
            int startColumn = 0;
            setValue(workRow, 0, storeCategorySales.getStoreName());
            setValue(workRow, 1, storeCategorySales.getSimilarStore());
            setValue(workRow, 2, storeCategorySales.getCategoryName());

            setValue(workRow, 3, storeCategorySales.getPrevPeriodCurrentMonthAvgTurnover());
            setValue(workRow, 4, storeCategorySales.getPrevPeriodCurrentMonthAvgMargin());
            setValue(workRow, 5, storeCategorySales.getPrevPeriodFirstMonthAvgTurnover());
            setValue(workRow, 6, storeCategorySales.getPrevPeriodFirstMonthAvgMargin());
            setValue(workRow, 7, storeCategorySales.getPrevPeriodSecondMonthAvgTurnover());
            setValue(workRow, 8, storeCategorySales.getPrevPeriodSecondMonthAvgMargin());
            setValue(workRow, 9, storeCategorySales.getPrevPeriodThirdMonthAvgTurnover());
            setValue(workRow, 10, storeCategorySales.getPrevPeriodThirdMonthAvgMargin());
            setValue(workRow, 11, storeCategorySales.getCurrentPeriodCurrentMonthAvgTurnover());
            setValue(workRow, 12, storeCategorySales.getCurrentPeriodCurrentMonthAvgMargin());

            setValue(workRow, 25, storeCategorySales.getPrevPeriodFirstMonthSalesDays());
            setValue(workRow, 26, storeCategorySales.getPrevPeriodSecondMonthSalesDays());
            setValue(workRow, 27, storeCategorySales.getPrevPeriodThirdMonthSalesDays());
        }
        return dataSheet;
    }

    private static void createDataSheetHeader(final Workbook workbook,
                                              final Quarter prevQuarter,
                                              final Quarter plannedQuarter,
                                              final Sheet dataSheet) {
        final Row firstRow = dataSheet.createRow(0);
        final Row secondRow = dataSheet.createRow(1);
        firstRow.createCell(0).setCellValue(STORE);
        firstRow.createCell(1).setCellValue(SIMILAR_STORE);
        firstRow.createCell(2).setCellValue(CATEGORY);

        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));

        firstRow.createCell(3).setCellValue(String.format("СДП %s", prevQuarter.actualMonthName()));
        firstRow.createCell(4);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 4));
        firstRow.createCell(5).setCellValue(String.format("СДП %s", prevQuarter.firstMonthName()));
        firstRow.createCell(6);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 5, 6));
        firstRow.createCell(7).setCellValue(String.format("СДП %s", prevQuarter.secondMonth()));
        firstRow.createCell(8);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 7, 8));
        firstRow.createCell(9).setCellValue(String.format("СДП %s", prevQuarter.thirdMonthName()));
        firstRow.createCell(10);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 9, 10));
        firstRow.createCell(11).setCellValue(String.format("СДП %s", plannedQuarter.actualMonthName()));
        firstRow.createCell(12);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 11, 12));

        firstRow.createCell(13).setCellValue(String.format("Дин %s / %s", prevQuarter.firstMonthName(), prevQuarter.actualMonthName()));
        firstRow.createCell(14);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 13, 14));
        firstRow.createCell(15).setCellValue(String.format("Дин %s / %s", prevQuarter.secondMonth(), prevQuarter.firstMonthName()));
        firstRow.createCell(16);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 15, 16));
        firstRow.createCell(17).setCellValue(String.format("Дин %s / %s", prevQuarter.thirdMonthName(), prevQuarter.secondMonth()));
        firstRow.createCell(18);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 17, 18));

        firstRow.createCell(19).setCellValue(String.format("План СДП %s", plannedQuarter.firstMonthName()));
        firstRow.createCell(20);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 19, 20));
        firstRow.createCell(21).setCellValue(String.format("План СДП %s", plannedQuarter.secondMonth()));
        firstRow.createCell(22);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 21, 22));
        firstRow.createCell(23).setCellValue(String.format("План СДП %s", plannedQuarter.thirdMonthName()));
        firstRow.createCell(24);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 23, 24));

        firstRow.createCell(25).setCellValue(String.format("Дні %s", plannedQuarter.firstMonthName()));
        firstRow.createCell(26).setCellValue(String.format("Дні %s", plannedQuarter.secondMonth()));
        firstRow.createCell(27).setCellValue(String.format("Дні %s", plannedQuarter.thirdMonthName()));

        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 25, 25));
        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 26, 26));
        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 27, 27));

        firstRow.createCell(28).setCellValue(String.format("План %s", plannedQuarter.firstMonthName()));
        firstRow.createCell(29);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 28, 29));
        firstRow.createCell(30).setCellValue(String.format("План %s", plannedQuarter.secondMonth()));
        firstRow.createCell(31);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 30, 31));
        firstRow.createCell(32).setCellValue(String.format("План %s", plannedQuarter.thirdMonthName()));
        firstRow.createCell(33);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 32, 33));

        IntStream.rangeClosed(0, 33).forEach(secondRow::createCell);

        for (int i = 3; i < 24; i = i + 2) {
            secondRow.getCell(i).setCellValue("ТО");
            secondRow.getCell(i + 1).setCellValue("Маржа");
        }

        for (int i = 28; i < 33; i = i + 2) {
            secondRow.getCell(i).setCellValue("ТО");
            secondRow.getCell(i + 1).setCellValue("Маржа");
        }

        final CellStyle centeredWrappedStyle = workbook.createCellStyle();
        centeredWrappedStyle.setAlignment(HorizontalAlignment.CENTER);
        centeredWrappedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        centeredWrappedStyle.setWrapText(true);

        applyCellStyle(dataSheet, centeredWrappedStyle, 0, 0, 1, 33);
    }

    private <T> void setValue(final Row row, final int col, final T value) {
        Cell cell = row.getCell(col);
        if (isNull(cell)) {
            cell = row.createCell(col);
        }
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
}
