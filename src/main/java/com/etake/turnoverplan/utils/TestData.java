package com.etake.turnoverplan.utils;

import com.etake.turnoverplan.model.RegionOrder;
import com.etake.turnoverplan.model.StoreCategorySales;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@NoArgsConstructor(access = AccessLevel.NONE)
public final class TestData {

    public static List<StoreCategorySales> getTestSales() {
        List<StoreCategorySales> storeCategorySalesList = new ArrayList<>();

        // Store 1 with 3 categories
        storeCategorySalesList.add(StoreCategorySales.builder()
                .regionName("Region1")
                .storeName("Store1")
                .similarStore("Store1")
                .categoryName("Category1")
                .currentPeriodCurrentMonthAvgTurnover(new BigDecimal("1000.00"))
                .prevPeriodCurrentMonthAvgTurnover(new BigDecimal("900.00"))
                .prevPeriodFirstMonthAvgTurnover(new BigDecimal("850.00"))
                .prevPeriodSecondMonthAvgTurnover(new BigDecimal("800.00"))
                .prevPeriodThirdMonthAvgTurnover(new BigDecimal("750.00"))
                .currentPeriodCurrentMonthAvgMargin(new BigDecimal("200.00"))
                .prevPeriodCurrentMonthAvgMargin(new BigDecimal("180.00"))
                .prevPeriodFirstMonthAvgMargin(new BigDecimal("170.00"))
                .prevPeriodSecondMonthAvgMargin(new BigDecimal("160.00"))
                .prevPeriodThirdMonthAvgMargin(new BigDecimal("150.00"))
                .prevPeriodFirstMonthSalesDays(30)
                .prevPeriodSecondMonthSalesDays(28)
                .prevPeriodThirdMonthSalesDays(31)
                .isLfl(true)
                .build());

        storeCategorySalesList.add(StoreCategorySales.builder()
                .regionName("Region1")
                .storeName("Store1")
                .similarStore("Store1")
                .categoryName("Category2")
                .currentPeriodCurrentMonthAvgTurnover(new BigDecimal("1100.00"))
                .prevPeriodCurrentMonthAvgTurnover(new BigDecimal("1000.00"))
                .prevPeriodFirstMonthAvgTurnover(new BigDecimal("950.00"))
                .prevPeriodSecondMonthAvgTurnover(new BigDecimal("900.00"))
                .prevPeriodThirdMonthAvgTurnover(new BigDecimal("850.00"))
                .currentPeriodCurrentMonthAvgMargin(new BigDecimal("220.00"))
                .prevPeriodCurrentMonthAvgMargin(new BigDecimal("200.00"))
                .prevPeriodFirstMonthAvgMargin(new BigDecimal("190.00"))
                .prevPeriodSecondMonthAvgMargin(new BigDecimal("180.00"))
                .prevPeriodThirdMonthAvgMargin(new BigDecimal("170.00"))
                .prevPeriodFirstMonthSalesDays(30)
                .prevPeriodSecondMonthSalesDays(28)
                .prevPeriodThirdMonthSalesDays(31)
                .isLfl(true)
                .build());

        storeCategorySalesList.add(StoreCategorySales.builder()
                .regionName("Region1")
                .storeName("Store1")
                .similarStore("Store1")
                .categoryName("Category3")
                .currentPeriodCurrentMonthAvgTurnover(new BigDecimal("1200.00"))
                .prevPeriodCurrentMonthAvgTurnover(new BigDecimal("1100.00"))
                .prevPeriodFirstMonthAvgTurnover(new BigDecimal("1050.00"))
                .prevPeriodSecondMonthAvgTurnover(new BigDecimal("1000.00"))
                .prevPeriodThirdMonthAvgTurnover(new BigDecimal("950.00"))
                .currentPeriodCurrentMonthAvgMargin(new BigDecimal("240.00"))
                .prevPeriodCurrentMonthAvgMargin(new BigDecimal("220.00"))
                .prevPeriodFirstMonthAvgMargin(new BigDecimal("210.00"))
                .prevPeriodSecondMonthAvgMargin(new BigDecimal("200.00"))
                .prevPeriodThirdMonthAvgMargin(new BigDecimal("190.00"))
                .prevPeriodFirstMonthSalesDays(30)
                .prevPeriodSecondMonthSalesDays(28)
                .prevPeriodThirdMonthSalesDays(31)
                .isLfl(true)
                .build());

        // Store 2 with 2 categories
        storeCategorySalesList.add(StoreCategorySales.builder()
                .regionName("Region2")
                .storeName("Store2")
                .similarStore("Store2")
                .categoryName("Category1")
                .currentPeriodCurrentMonthAvgTurnover(new BigDecimal("1300.00"))
                .prevPeriodCurrentMonthAvgTurnover(new BigDecimal("1200.00"))
                .prevPeriodFirstMonthAvgTurnover(new BigDecimal("1150.00"))
                .prevPeriodSecondMonthAvgTurnover(new BigDecimal("1100.00"))
                .prevPeriodThirdMonthAvgTurnover(new BigDecimal("1050.00"))
                .currentPeriodCurrentMonthAvgMargin(new BigDecimal("260.00"))
                .prevPeriodCurrentMonthAvgMargin(new BigDecimal("240.00"))
                .prevPeriodFirstMonthAvgMargin(new BigDecimal("230.00"))
                .prevPeriodSecondMonthAvgMargin(new BigDecimal("220.00"))
                .prevPeriodThirdMonthAvgMargin(new BigDecimal("210.00"))
                .prevPeriodFirstMonthSalesDays(30)
                .prevPeriodSecondMonthSalesDays(28)
                .prevPeriodThirdMonthSalesDays(31)
                .isLfl(true)
                .build());

        storeCategorySalesList.add(StoreCategorySales.builder()
                .regionName("Region2")
                .storeName("Store2")
                .similarStore("Store2")
                .categoryName("Category2")
                .currentPeriodCurrentMonthAvgTurnover(new BigDecimal("1400.00"))
                .prevPeriodCurrentMonthAvgTurnover(new BigDecimal("1300.00"))
                .prevPeriodFirstMonthAvgTurnover(new BigDecimal("1250.00"))
                .prevPeriodSecondMonthAvgTurnover(new BigDecimal("1200.00"))
                .prevPeriodThirdMonthAvgTurnover(new BigDecimal("1150.00"))
                .currentPeriodCurrentMonthAvgMargin(new BigDecimal("280.00"))
                .prevPeriodCurrentMonthAvgMargin(new BigDecimal("260.00"))
                .prevPeriodFirstMonthAvgMargin(new BigDecimal("250.00"))
                .prevPeriodSecondMonthAvgMargin(new BigDecimal("240.00"))
                .prevPeriodThirdMonthAvgMargin(new BigDecimal("230.00"))
                .prevPeriodFirstMonthSalesDays(30)
                .prevPeriodSecondMonthSalesDays(28)
                .prevPeriodThirdMonthSalesDays(31)
                .isLfl(true)
                .build());

        // Store 3 with 1 category
        storeCategorySalesList.add(StoreCategorySales.builder()
                .regionName("Region3")
                .storeName("Store3")
                .similarStore("Store3")
                .categoryName("Category1")
                .currentPeriodCurrentMonthAvgTurnover(new BigDecimal("1500.00"))
                .prevPeriodCurrentMonthAvgTurnover(new BigDecimal("1400.00"))
                .prevPeriodFirstMonthAvgTurnover(new BigDecimal("1350.00"))
                .prevPeriodSecondMonthAvgTurnover(new BigDecimal("1300.00"))
                .prevPeriodThirdMonthAvgTurnover(new BigDecimal("1250.00"))
                .currentPeriodCurrentMonthAvgMargin(new BigDecimal("300.00"))
                .prevPeriodCurrentMonthAvgMargin(new BigDecimal("280.00"))
                .prevPeriodFirstMonthAvgMargin(new BigDecimal("270.00"))
                .prevPeriodSecondMonthAvgMargin(new BigDecimal("260.00"))
                .prevPeriodThirdMonthAvgMargin(new BigDecimal("250.00"))
                .prevPeriodFirstMonthSalesDays(30)
                .prevPeriodSecondMonthSalesDays(28)
                .prevPeriodThirdMonthSalesDays(31)
                .isLfl(true)
                .build());

        // Populate similarStore field
        Map<String, List<StoreCategorySales>> storesByRegion = storeCategorySalesList.stream()
                .collect(Collectors.groupingBy(StoreCategorySales::getRegionName));

        for (StoreCategorySales store : storeCategorySalesList) {
            List<StoreCategorySales> sameRegionStores = storesByRegion.get(store.getRegionName());
            if (sameRegionStores.size() > 1) {
                for (StoreCategorySales similarStore : sameRegionStores) {
                    if (!similarStore.getStoreName().equals(store.getStoreName())) {
                        store.setSimilarStore(similarStore.getStoreName());
                        break;
                    }
                }
            }
        }

        return storeCategorySalesList;
    }

    public static Collection<RegionOrder> getRegionOrders() {
        return List.of(new RegionOrder(1, "Region1"), new RegionOrder(2, "Region2"), new RegionOrder(3, "Region3"));
    }
}
