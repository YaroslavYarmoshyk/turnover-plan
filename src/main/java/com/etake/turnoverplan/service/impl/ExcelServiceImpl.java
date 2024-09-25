package com.etake.turnoverplan.service.impl;

import com.etake.turnoverplan.model.Quarter;
import com.etake.turnoverplan.model.StoreCategorySales;
import com.etake.turnoverplan.service.ExcelService;
import com.etake.turnoverplan.service.PeriodProvider;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collection;
import java.util.stream.IntStream;

import static com.etake.turnoverplan.utils.Constants.CATEGORY;
import static com.etake.turnoverplan.utils.Constants.SIMILAR_STORE;
import static com.etake.turnoverplan.utils.Constants.STORE;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {
    private final PeriodProvider periodProvider;

    @Override
    public Workbook getWorkbook(final Collection<StoreCategorySales> sales) {
        try (final Workbook workbook = new SXSSFWorkbook()) {
            final Quarter prevQuarter = periodProvider.getPrevQuarter();
            final Quarter plannedQuarter = periodProvider.getPlannedQuarter();
            final Sheet dataSheet = getDataSheet(workbook, prevQuarter, plannedQuarter);
            final Sheet storesPlan = workbook.createSheet("План по магазинам");
            final Sheet categoriesPlan = workbook.createSheet("План по категоріям");

            return workbook;
        } catch (final Exception e) {
            throw new IllegalStateException("Cannot create report workbook");
        }
    }

    @Override
    public void writeWorkbook(final Workbook workbook, final String path) {
        try (final FileOutputStream fileOut = new FileOutputStream(path)) {
            workbook.write(fileOut);
            workbook.close();
        } catch (IOException e) {
            throw new IllegalStateException("Cannot write workbook to file", e);
        }
    }

    private static Sheet getDataSheet(final Workbook workbook, final Quarter prevQuarter, final Quarter plannedQuarter) {
        final Sheet dataSheet = workbook.createSheet("data");
        final Row firstRow = dataSheet.createRow(0);
        firstRow.createCell(0).setCellValue(STORE);
        firstRow.createCell(1).setCellValue(SIMILAR_STORE);
        firstRow.createCell(2).setCellValue(CATEGORY);

        firstRow.createCell(3).setCellValue(String.format("СДП %s", prevQuarter.actualMonthName()));
        firstRow.createCell(4);
        firstRow.createCell(5).setCellValue(String.format("СДП %s", prevQuarter.firstMonthName()));
        firstRow.createCell(6);
        firstRow.createCell(7).setCellValue(String.format("СДП %s", prevQuarter.secondMonth()));
        firstRow.createCell(8);
        firstRow.createCell(9).setCellValue(String.format("СДП %s", prevQuarter.thirdMonthName()));
        firstRow.createCell(10);
        firstRow.createCell(11).setCellValue(String.format("СДП %s", plannedQuarter.actualMonthName()));
        firstRow.createCell(12);

        firstRow.createCell(13).setCellValue(String.format("Дин %s / %s", prevQuarter.firstMonthName(), prevQuarter.actualMonthName()));
        firstRow.createCell(14);
        firstRow.createCell(15).setCellValue(String.format("Дин %s / %s", prevQuarter.secondMonth(), prevQuarter.firstMonthName()));
        firstRow.createCell(16);
        firstRow.createCell(17).setCellValue(String.format("Дин %s / %s", prevQuarter.thirdMonthName(), prevQuarter.secondMonth()));
        firstRow.createCell(18);

        firstRow.createCell(19).setCellValue(String.format("План СДП %s", plannedQuarter.firstMonthName()));
        firstRow.createCell(20);
        firstRow.createCell(21).setCellValue(String.format("План СДП %s", plannedQuarter.secondMonth()));
        firstRow.createCell(22);
        firstRow.createCell(23).setCellValue(String.format("План СДП %s", plannedQuarter.thirdMonthName()));
        firstRow.createCell(24);

        firstRow.createCell(25).setCellValue(String.format("Дні %s", plannedQuarter.firstMonthName()));
        firstRow.createCell(26).setCellValue(String.format("Дні %s", plannedQuarter.secondMonth()));
        firstRow.createCell(27).setCellValue(String.format("Дні %s", plannedQuarter.thirdMonthName()));

        firstRow.createCell(28).setCellValue(String.format("План %s", plannedQuarter.firstMonthName()));
        firstRow.createCell(29);
        firstRow.createCell(30).setCellValue(String.format("План %s", plannedQuarter.secondMonth()));
        firstRow.createCell(31);
        firstRow.createCell(32).setCellValue(String.format("План %s", plannedQuarter.thirdMonthName()));
        firstRow.createCell(33);

        final Row secondRow = dataSheet.createRow(1);
        IntStream.rangeClosed(0, 33).forEach(secondRow::createCell);
        IntStream.range(3, 24).forEach(i -> {
            secondRow.getCell(i).setCellValue("ТО");
            secondRow.getCell(i + 1).setCellValue("Маржа");
        });

        IntStream.range(28, 33).forEach(i -> {
            secondRow.getCell(i).setCellValue("ТО");
            secondRow.getCell(i + 1).setCellValue("Маржа");
        });

        return dataSheet;
    }
}
