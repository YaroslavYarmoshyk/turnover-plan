package com.etake.turnoverplan.service.impl;

import com.etake.turnoverplan.model.Position;
import com.etake.turnoverplan.model.StoreCategory;
import com.etake.turnoverplan.model.StoreCategorySales;
import com.etake.turnoverplan.repository.PositionRepository;
import com.etake.turnoverplan.repository.StoreCategoryRepository;
import com.etake.turnoverplan.service.StoreCategoryService;
import lombok.RequiredArgsConstructor;
import org.apache.logging.log4j.util.Strings;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.temporal.TemporalAdjusters;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

import static com.etake.turnoverplan.utils.CalculationUtils.DEFAULT_ROUNDING_MODE;
import static java.util.stream.Collectors.groupingBy;
import static java.util.stream.Collectors.toMap;
import static java.util.stream.Collectors.toSet;

@Service
@RequiredArgsConstructor
public class StoreCategoryServiceImpl implements StoreCategoryService {
    private final StoreCategoryRepository storeCategoryRepository;
    private final PositionRepository positionRepository;

    @Override
    public List<StoreCategorySales> getSales(final Integer year, final Integer month) {
        final YearMonth currentPeriodCurrentMonth = YearMonth.of(year, month);
        final YearMonth currentPeriodPrevMonth = getPrevPeriodYearMonth(currentPeriodCurrentMonth, 0, -1);
        final YearMonth prevPeriodPrevMonth = getPrevPeriodYearMonth(currentPeriodCurrentMonth, -1, -1);
        final YearMonth prevPeriodCurrentMonth = getPrevPeriodYearMonth(currentPeriodCurrentMonth, -1, 0);
        final YearMonth prevPeriodFirstMonth = getPrevPeriodYearMonth(currentPeriodCurrentMonth, -1, 1);
        final YearMonth prevPeriodSecondMonth = getPrevPeriodYearMonth(currentPeriodCurrentMonth, -1, 2);
        final YearMonth prevPeriodThirdMonth = getPrevPeriodYearMonth(currentPeriodCurrentMonth, -1, 3);

        final Map<String, Position> currentPeriodPrevMonthPositions = getPositionsByKey(currentPeriodPrevMonth);
        final Map<String, Position> currentPeriodCurrentMonthPositions = getPositionsByKey(currentPeriodCurrentMonth);
        final Map<String, Position> prevPeriodCurrentMonthPositions = getPositionsByKey(prevPeriodCurrentMonth);
        final Map<String, Position> prevPeriodFirstMonthPositions = getPositionsByKey(prevPeriodFirstMonth);
        final Map<String, Position> prevPeriodSecondMonthPositions = getPositionsByKey(prevPeriodSecondMonth);
        final Map<String, Position> prevPeriodThirdMonthPositions = getPositionsByKey(prevPeriodThirdMonth);

        final List<String> lflStores = getPositionsByKey(prevPeriodPrevMonth).values().stream()
                .map(Position::storeName)
                .toList();

        final List<StoreCategory> activeStoreCategories = storeCategoryRepository.getActiveStoreCategoryPairs();
        final Map<String, Map<String, Set<StoreCategory>>> activeStoreCategoriesByRegion = activeStoreCategories.stream()
                .collect(groupingBy(
                        StoreCategory::regionName,
                        groupingBy(
                                StoreCategory::categoryName,
                                toSet()
                        )
                ));

        final List<StoreCategorySales> storeCategorySales = activeStoreCategories.stream()
                .map(storeCategory -> {
                    final String regionName = storeCategory.regionName();
                    final String storeName = storeCategory.storeName();
                    final String categoryName = storeCategory.categoryName();
                    final String key = storeCategory.getKey();
                    final boolean isLfl = lflStores.contains(storeName);

                    final Set<StoreCategory> storeCategoriesInRegion = activeStoreCategoriesByRegion.get(regionName).get(categoryName);
                    final String similarStoreName = isLfl ? storeName : resolveSimilarStore(
                            storeCategory,
                            storeCategoriesInRegion,
                            getFilteredByRegionCategoryPositions(storeCategoriesInRegion, currentPeriodPrevMonthPositions),
                            getFilteredByRegionCategoryPositions(storeCategoriesInRegion, currentPeriodCurrentMonthPositions)
                    );
                    return StoreCategorySales.builder()
                            .regionName(regionName)
                            .storeName(storeName)
                            .categoryName(storeCategory.categoryName())

                            .currentPeriodCurrentMonthAvgTurnover(getValueByKey(key, currentPeriodCurrentMonthPositions, Position::turnover))
                            .prevPeriodCurrentMonthAvgTurnover(getValueByKey(key, prevPeriodCurrentMonthPositions, Position::turnover))
                            .prevPeriodFirstMonthAvgTurnover(getValueByKey(key, prevPeriodFirstMonthPositions, Position::turnover))
                            .prevPeriodSecondMonthAvgTurnover(getValueByKey(key, prevPeriodSecondMonthPositions, Position::turnover))
                            .prevPeriodThirdMonthAvgTurnover(getValueByKey(key, prevPeriodThirdMonthPositions, Position::turnover))

                            .currentPeriodCurrentMonthAvgMargin(getValueByKey(key, currentPeriodCurrentMonthPositions, Position::margin))
                            .prevPeriodCurrentMonthAvgMargin(getValueByKey(key, prevPeriodCurrentMonthPositions, Position::margin))
                            .prevPeriodFirstMonthAvgMargin(getValueByKey(key, prevPeriodFirstMonthPositions, Position::margin))
                            .prevPeriodSecondMonthAvgMargin(getValueByKey(key, prevPeriodSecondMonthPositions, Position::margin))
                            .prevPeriodThirdMonthAvgMargin(getValueByKey(key, prevPeriodThirdMonthPositions, Position::margin))

                            .prevPeriodFirstMonthSalesDays(getValueByKey(key, prevPeriodFirstMonthPositions, Position::salesDays))
                            .prevPeriodSecondMonthSalesDays(getValueByKey(key, prevPeriodSecondMonthPositions, Position::salesDays))
                            .prevPeriodThirdMonthSalesDays(getValueByKey(key, prevPeriodThirdMonthPositions, Position::salesDays))

                            .isLfl(isLfl)

                            .similarStore(similarStoreName)

                            .build();
                })
                .collect(Collectors.toList());

        return List.of();
    }

    private String resolveSimilarStore(final StoreCategory storeCategory,
                                       final Set<StoreCategory> storeCategories,
                                       final Map<String, Position> currentPeriodPrevMonthPositions,
                                       final Map<String, Position> currentPeriodCurrentMonthPositions) {
        final Map<String, BigDecimal> growth = currentPeriodCurrentMonthPositions.entrySet().stream()
                .collect(toMap(
                                Map.Entry::getKey,
                                entry -> {
                                    final BigDecimal currentTurnover = entry.getValue().turnover();
                                    final BigDecimal prevTurnover = currentPeriodPrevMonthPositions.getOrDefault(entry.getKey(), Position.empty()).turnover();
//                                    If new store, return zero growth
                                    if (currentTurnover == null || prevTurnover == null) {
                                        return BigDecimal.ZERO;
                                    }
                                    return currentTurnover
                                            .divide(prevTurnover, DEFAULT_ROUNDING_MODE).subtract(BigDecimal.ONE);
                                }
                        )
                );
        final String currentStoreName = storeCategory.storeName();
        final String currentStoreCategoryKey = storeCategory.getKey();
        final BigDecimal currentStoreCategoryGrowth = growth.get(currentStoreCategoryKey);
//        If a category is not present on store, return the same stores as similar one
        if (currentStoreCategoryGrowth == null) {
            return currentStoreName;
        }
        final Map<String, BigDecimal> growthExcludingCurrentStoreCategory = growth.entrySet().stream()
                .filter(entry -> !Objects.equals(entry.getKey(), currentStoreCategoryKey))
                .collect(toMap(Map.Entry::getKey, Map.Entry::getValue));
        final String similarStoreCategoryKey = getClosestStoreCategoryKey(currentStoreCategoryGrowth, growthExcludingCurrentStoreCategory);
        if (Strings.isEmpty(similarStoreCategoryKey)) {
            return currentStoreName;
        }

        return storeCategories.stream()
                .filter(sC -> Objects.equals(sC.getKey(), similarStoreCategoryKey))
                .findFirst()
                .map(StoreCategory::storeName)
                .orElseThrow();
    }

    private static Map<String, Position> getFilteredByRegionCategoryPositions(final Collection<StoreCategory> storeCategories,
                                                                              final Map<String, Position> positions) {
        final String categoryName = storeCategories.stream()
//                .map(StoreCategory::categoryName)
                .map(s -> {
                    if (s.categoryName() == null) {
                        String stop = "s";
                    }
                    return s.categoryName();
                })
                .findFirst()
                .orElseThrow();
        final Set<String> storesInRegion = storeCategories.stream()
                .map(StoreCategory::storeName)
                .collect(toSet());
        return positions.entrySet().stream()
                .filter(entry -> Objects.equals(entry.getValue().categoryName(), categoryName) && storesInRegion.contains(entry.getValue().storeName()))
                .collect(toMap(Map.Entry::getKey, Map.Entry::getValue));
    }

    private static String getClosestStoreCategoryKey(final BigDecimal target, final Map<String, BigDecimal> growth) {
        return growth.entrySet().stream()
                .filter(entry -> entry.getValue() != null)  // Filter out entries with null values
                .min((entry1, entry2) -> {
                    BigDecimal diff1 = entry1.getValue().subtract(target).abs();
                    BigDecimal diff2 = entry2.getValue().subtract(target).abs();
                    return diff1.compareTo(diff2);  // Compare the absolute differences
                })
                .map(Map.Entry::getKey)
                .orElse(Strings.EMPTY);
    }

    private <T> T getValueByKey(final String key,
                                final Map<String, Position> positions,
                                final Function<Position, T> function) {
        final Position position = positions.getOrDefault(key, Position.empty());
        return function.apply(position);
    }

    private Map<String, Position> getPositionsByKey(final YearMonth yearMonth) {
        final LocalDate fromDate = getStartDate(yearMonth);
        final LocalDate toDay = getEndDate(yearMonth);
        final List<Position> positions = positionRepository.findAllInPeriod(fromDate, toDay);
        return positions.stream()
                .collect(toMap(
                        Position::getKey,
                        Function.identity()
                ));
    }

    private static YearMonth getPrevPeriodYearMonth(final YearMonth currentYearMonth,
                                                    final Integer yearOffset,
                                                    final Integer monthOffset) {
        return currentYearMonth.plusYears(yearOffset).plusMonths(monthOffset);
    }

    private static LocalDate getStartDate(final YearMonth yearMonth) {
        return LocalDate.of(yearMonth.getYear(), yearMonth.getMonth(), 1);
    }

    private static LocalDate getEndDate(final YearMonth yearMonth) {
        return getStartDate(yearMonth).with(TemporalAdjusters.lastDayOfMonth());
    }
}
