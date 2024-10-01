package com.etake.turnoverplan.service;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelFormatService {

    void formatDataSheet(final Workbook workbook, final Sheet sheet);

    void formatCategoriesSheet(final Workbook workbook, final Sheet sheet);

    void formatStoresSheet(final Workbook workbook, final Sheet sheet);
}