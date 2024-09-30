package com.etake.turnoverplan.config.indices;

import java.util.Set;

public record DataSheetColumnIndices(
        int store,
        int similarStore,
        int category,
        int avgTurnoverPrevPeriodActualMonth,
        int avgMarginPrevPeriodActualMonth,
        int avgTurnoverPrevPeriodFirstMonth,
        int avgMarginPrevPeriodFirstMonth,
        int avgTurnoverPrevPeriodSecondMonth,
        int avgMarginPrevPeriodSecondMonth,
        int avgTurnoverPrevPeriodThirdMonth,
        int avgMarginPrevPeriodThirdMonth,
        int avgTurnoverCurrentPeriodActualMonth,
        int avgMarginCurrentPeriodActualMonth,
        int dynTurnoverFirstToActualMonth,
        int dynMarginFirstToActualMonth,
        int dynTurnoverSecondToFirstMonth,
        int dynMarginSecondToFirstMonth,
        int dynTurnoverThirdToSecondMonth,
        int dynMarginThirdToSecondMonth,
        int planAvgTurnoverFirstMonth,
        int planAvgMarginFirstMonth,
        int planAvgTurnoverSecondMonth,
        int planAvgMarginSecondMonth,
        int planAvgTurnoverThirdMonth,
        int planAvgMarginThirdMonth,
        int daysFirstMonth,
        int daysSecondMonth,
        int daysThirdMonth,
        int planTurnoverFirstMonth,
        int planMarginFirstMonth,
        int planTurnoverSecondMonth,
        int planMarginSecondMonth,
        int planTurnoverThirdMonth,
        int planMarginThirdMonth
) {
    public Set<Integer> getTurnoverColumnIndices() {
        return Set.of(
            avgTurnoverPrevPeriodActualMonth,
            avgTurnoverPrevPeriodFirstMonth,
            avgTurnoverPrevPeriodSecondMonth,
            avgTurnoverPrevPeriodThirdMonth,
            avgTurnoverCurrentPeriodActualMonth,
            dynTurnoverFirstToActualMonth,
            dynTurnoverSecondToFirstMonth,
            dynTurnoverThirdToSecondMonth,
            planAvgTurnoverFirstMonth,
            planAvgTurnoverSecondMonth,
            planAvgTurnoverThirdMonth,
            planTurnoverFirstMonth,
            planTurnoverSecondMonth,
            planTurnoverThirdMonth
        );
    }

    public Set<Integer> getMarginColumnIndices() {
        return Set.of(
                avgMarginPrevPeriodActualMonth,
                avgMarginPrevPeriodFirstMonth,
                avgMarginPrevPeriodSecondMonth,
                avgMarginPrevPeriodThirdMonth,
                avgMarginCurrentPeriodActualMonth,
                dynMarginFirstToActualMonth,
                dynMarginSecondToFirstMonth,
                dynMarginThirdToSecondMonth,
                planAvgMarginFirstMonth,
                planAvgMarginSecondMonth,
                planAvgMarginThirdMonth,
                planMarginFirstMonth,
                planMarginSecondMonth,
                planMarginThirdMonth
        );
    }
}
