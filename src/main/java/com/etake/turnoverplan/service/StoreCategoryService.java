package com.etake.turnoverplan.service;

import com.etake.turnoverplan.model.StoreCategorySales;

import java.util.List;

public interface StoreCategoryService {

    List<StoreCategorySales> getSales(final Integer year, final Integer month);
}
