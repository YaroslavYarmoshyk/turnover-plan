package com.etake.turnoverplan.service.impl;

import com.etake.turnoverplan.model.Quarter;
import com.etake.turnoverplan.model.StoreCategorySales;
import com.etake.turnoverplan.service.ExcelService;
import com.etake.turnoverplan.service.PeriodProvider;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.util.Collection;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static com.etake.turnoverplan.utils.Constants.CATEGORY;
import static com.etake.turnoverplan.utils.Constants.SIMILAR_STORE;
import static com.etake.turnoverplan.utils.Constants.STORE;
import static com.etake.turnoverplan.utils.ExcelFormulaUtils.getAvgPlanFormula;
import static com.etake.turnoverplan.utils.ExcelFormulaUtils.getCategoriesPlanFormula;
import static com.etake.turnoverplan.utils.ExcelFormulaUtils.getDynFormula;
import static com.etake.turnoverplan.utils.ExcelFormulaUtils.getPlanFormula;
import static com.etake.turnoverplan.utils.ExcelFormulaUtils.getTotalSumFormula;
import static com.etake.turnoverplan.utils.ExcelFormulaUtils.setFormula;
import static com.etake.turnoverplan.utils.ExcelUtils.applyCellStyle;
import static com.etake.turnoverplan.utils.ExcelUtils.createCells;
import static com.etake.turnoverplan.utils.ExcelUtils.setValue;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {
    private static final Integer INITIAL_VALUE_ROW_INDEX = 2;
    private final PeriodProvider periodProvider;

    @Override
    public Workbook getWorkbook(final Collection<StoreCategorySales> sales) {
        final Workbook workbook = new XSSFWorkbook();
        final Quarter prevQuarter = periodProvider.getPrevQuarter();
        final Quarter plannedQuarter = periodProvider.getPlannedQuarter();
        final Sheet dataSheet = getDataSheet(workbook, prevQuarter, plannedQuarter, sales);
        final Sheet categoriesPlan = getCategoriesSheet(workbook, plannedQuarter, sales);
        final Sheet storesPlan = workbook.createSheet("План по магазинам");

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

        createCells(dataSheet, INITIAL_VALUE_ROW_INDEX + sales.size(), 33);

        createDataSheetHeaders(workbook, dataSheet, prevQuarter, plannedQuarter);
        populateCalculatedData(dataSheet, sales);
        populateFormulas(dataSheet, sales.size());

        return dataSheet;
    }

    private static void createDataSheetHeaders(final Workbook workbook, final Sheet dataSheet,
                                               final Quarter prevQuarter, final Quarter plannedQuarter) {
        final Row firstRow = dataSheet.getRow(0);
        final Row secondRow = dataSheet.getRow(1);
        firstRow.getCell(0).setCellValue(STORE);
        firstRow.getCell(1).setCellValue(SIMILAR_STORE);
        firstRow.getCell(2).setCellValue(CATEGORY);

        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));

        firstRow.getCell(3).setCellValue(String.format("СДП %s", prevQuarter.actualMonthName()));
        firstRow.getCell(4);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 4));
        firstRow.getCell(5).setCellValue(String.format("СДП %s", prevQuarter.firstMonthName()));
        firstRow.getCell(6);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 5, 6));
        firstRow.getCell(7).setCellValue(String.format("СДП %s", prevQuarter.secondMonth()));
        firstRow.getCell(8);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 7, 8));
        firstRow.getCell(9).setCellValue(String.format("СДП %s", prevQuarter.thirdMonthName()));
        firstRow.getCell(10);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 9, 10));
        firstRow.getCell(11).setCellValue(String.format("СДП %s", plannedQuarter.actualMonthName()));
        firstRow.getCell(12);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 11, 12));

        firstRow.getCell(13).setCellValue(String.format("Дин %s / %s", prevQuarter.firstMonthName(), prevQuarter.actualMonthName()));
        firstRow.getCell(14);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 13, 14));
        firstRow.getCell(15).setCellValue(String.format("Дин %s / %s", prevQuarter.secondMonth(), prevQuarter.firstMonthName()));
        firstRow.getCell(16);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 15, 16));
        firstRow.getCell(17).setCellValue(String.format("Дин %s / %s", prevQuarter.thirdMonthName(), prevQuarter.secondMonth()));
        firstRow.getCell(18);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 17, 18));

        firstRow.getCell(19).setCellValue(String.format("План СДП %s", plannedQuarter.firstMonthName()));
        firstRow.getCell(20);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 19, 20));
        firstRow.getCell(21).setCellValue(String.format("План СДП %s", plannedQuarter.secondMonth()));
        firstRow.getCell(22);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 21, 22));
        firstRow.getCell(23).setCellValue(String.format("План СДП %s", plannedQuarter.thirdMonthName()));
        firstRow.getCell(24);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 23, 24));

        firstRow.getCell(25).setCellValue(String.format("Дні %s", plannedQuarter.firstMonthName()));
        firstRow.getCell(26).setCellValue(String.format("Дні %s", plannedQuarter.secondMonth()));
        firstRow.getCell(27).setCellValue(String.format("Дні %s", plannedQuarter.thirdMonthName()));

        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 25, 25));
        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 26, 26));
        dataSheet.addMergedRegion(new CellRangeAddress(0, 1, 27, 27));

        firstRow.getCell(28).setCellValue(String.format("План %s", plannedQuarter.firstMonthName()));
        firstRow.getCell(29);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 28, 29));
        firstRow.getCell(30).setCellValue(String.format("План %s", plannedQuarter.secondMonth()));
        firstRow.getCell(31);
        dataSheet.addMergedRegion(new CellRangeAddress(0, 0, 30, 31));
        firstRow.getCell(32).setCellValue(String.format("План %s", plannedQuarter.thirdMonthName()));
        firstRow.getCell(33);
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

    private static void populateCalculatedData(final Sheet dataSheet, final Collection<StoreCategorySales> sales) {
        int startRow = 2;
        for (final StoreCategorySales storeCategorySales : sales) {
            final Row workRow = dataSheet.getRow(startRow++);
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
    }

    private static void populateFormulas(final Sheet dataSheet, final Integer positionsSize) {
        final Integer endRow = INITIAL_VALUE_ROW_INDEX + positionsSize;
        setFormula(getDynFormula("F:F", "D:D"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 13);
        setFormula(getDynFormula("G:G", "E:E"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 14);
        setFormula(getDynFormula("H:H", "F:F"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 15);
        setFormula(getDynFormula("I:I", "G:G"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 16);
        setFormula(getDynFormula("J:J", "H:H"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 17);
        setFormula(getDynFormula("K:K", "I:I"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 18);

        setFormula(getAvgPlanFormula("L", "N"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 19);
        setFormula(getAvgPlanFormula("M", "O"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 20);
        setFormula(getAvgPlanFormula("T", "P"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 21);
        setFormula(getAvgPlanFormula("U", "Q"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 22);
        setFormula(getAvgPlanFormula("V", "R"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 23);
        setFormula(getAvgPlanFormula("W", "S"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 24);

        setFormula(getPlanFormula("T", "Z:Z"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 28);
        setFormula(getPlanFormula("U", "Z:Z"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 29);
        setFormula(getPlanFormula("V", "AA:AA"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 30);
        setFormula(getPlanFormula("W", "AA:AA"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 31);
        setFormula(getPlanFormula("X", "AB:AB"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 32);
        setFormula(getPlanFormula("Y", "AB:AB"), dataSheet, INITIAL_VALUE_ROW_INDEX, endRow, 33);
    }

    private Sheet getCategoriesSheet(final Workbook workbook,
                                     final Quarter plannedQuarter,
                                     final Collection<StoreCategorySales> sales) {
        final Sheet categoriesSheet = workbook.createSheet("План по категоріям");
        final Set<String> categories = sales.stream()
                .map(StoreCategorySales::getCategoryName)
                .collect(Collectors.toSet());
        final int endRow = INITIAL_VALUE_ROW_INDEX + categories.size();

        createCells(categoriesSheet, endRow, 6);
        createCategoriesSheetHeaders(categoriesSheet, plannedQuarter);
        populateCategoriesSheetFormulas(categoriesSheet, endRow, categories);

        return categoriesSheet;
    }

    private static void populateCategoriesSheetFormulas(final Sheet categoriesSheet, final int endRow, final Set<String> categories) {
        int startRow = INITIAL_VALUE_ROW_INDEX;
        for (final String category : categories) {
            categoriesSheet.getRow(startRow++).getCell(0).setCellValue(category);
        }

        setFormula(getCategoriesPlanFormula("AC:AC"), categoriesSheet, INITIAL_VALUE_ROW_INDEX, endRow, 1);
        setFormula(getCategoriesPlanFormula("AD:AD"), categoriesSheet, INITIAL_VALUE_ROW_INDEX, endRow, 2);
        setFormula(getCategoriesPlanFormula("AE:AE"), categoriesSheet, INITIAL_VALUE_ROW_INDEX, endRow, 3);
        setFormula(getCategoriesPlanFormula("AF:AF"), categoriesSheet, INITIAL_VALUE_ROW_INDEX, endRow, 4);
        setFormula(getCategoriesPlanFormula("AG:AG"), categoriesSheet, INITIAL_VALUE_ROW_INDEX, endRow, 5);
        setFormula(getCategoriesPlanFormula("AH:AH"), categoriesSheet, INITIAL_VALUE_ROW_INDEX, endRow, 6);

        categoriesSheet.getRow(endRow).getCell(0).setCellValue("Разом");
        setFormula(getTotalSumFormula(INITIAL_VALUE_ROW_INDEX + 1, endRow), categoriesSheet, endRow, endRow, 1, 6);
    }

    private static void createCategoriesSheetHeaders(final Sheet categoriesSheet,
                                                     final Quarter plannedQuarter) {
        final Row firstRow = categoriesSheet.getRow(0);
        final Row secondRow = categoriesSheet.getRow(1);
        firstRow.getCell(0).setCellValue(STORE);
        categoriesSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));

        firstRow.getCell(1).setCellValue(String.format("План %s", plannedQuarter.firstMonthName()));
        firstRow.getCell(2);
        categoriesSheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 2));
        firstRow.getCell(3).setCellValue(String.format("План %s", plannedQuarter.secondMonth()));
        firstRow.getCell(4);
        categoriesSheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 4));
        firstRow.getCell(5).setCellValue(String.format("План %s", plannedQuarter.thirdMonthName()));
        firstRow.getCell(6);
        categoriesSheet.addMergedRegion(new CellRangeAddress(0, 0, 5, 6));

        for (int i = 1; i < 6; i = i + 2) {
            secondRow.getCell(i).setCellValue("ТО");
            secondRow.getCell(i + 1).setCellValue("Маржа");
        }
    }
}
