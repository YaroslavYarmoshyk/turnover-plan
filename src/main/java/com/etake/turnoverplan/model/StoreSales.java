package com.etake.turnoverplan.model;

import lombok.Data;
import lombok.experimental.SuperBuilder;

import java.math.BigDecimal;

@Data
@SuperBuilder
public class StoreSales {
    private String regionName;
    private String storeName;
    private BigDecimal currentPeriodCurrentMonthAvgTurnover;
    private BigDecimal prevPeriodCurrentMonthAvgTurnover;
    private BigDecimal prevPeriodFirstMonthAvgTurnover;
    private BigDecimal prevPeriodSecondMonthAvgTurnover;
    private BigDecimal prevPeriodThirdMonthAvgTurnover;

    private BigDecimal currentPeriodCurrentMonthAvgMargin;
    private BigDecimal prevPeriodCurrentMonthAvgMargin;
    private BigDecimal prevPeriodFirstMonthAvgMargin;
    private BigDecimal prevPeriodSecondMonthAvgMargin;
    private BigDecimal prevPeriodThirdMonthAvgMargin;

    private Integer prevPeriodFirstMonthSalesDays;
    private Integer prevPeriodSecondMonthSalesDays;
    private Integer prevPeriodThirdMonthSalesDays;

    private Boolean isLfl;
}
