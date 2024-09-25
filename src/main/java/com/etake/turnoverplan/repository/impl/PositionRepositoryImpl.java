package com.etake.turnoverplan.repository.impl;

import com.etake.turnoverplan.model.Position;
import com.etake.turnoverplan.repository.PositionRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.jdbc.core.simple.JdbcClient;
import org.springframework.stereotype.Repository;

import java.time.LocalDate;
import java.util.List;

@Repository
@RequiredArgsConstructor
public class PositionRepositoryImpl implements PositionRepository {
    private final JdbcClient jdbcClient;

    @Override
    public List<Position> findAllInPeriod(final LocalDate fromDate, final LocalDate toDate) {
        return jdbcClient.sql(getFindAllByYearMonthSqlQuery())
                .param("fromDate", fromDate)
                .param("toDate", toDate)
                .query(Position.class)
                .list();
    }

    private static String getFindAllByYearMonthSqlQuery() {
        return """
                SELECT
                    grouped.store_name,
                    grouped.category_name,
                    AVG(turnover) AS turnover,
                    AVG(margin) AS margin,
                    COUNT(date) AS sales_days
                FROM (
                    SELECT
                        s.date AS date,
                        st.name AS store_name,
                        cc.category_name,
                        SUM(s.discounted_sales) AS turnover,
                        SUM(s.margin) AS margin
                    FROM
                        stores st
                        LEFT JOIN sales s ON st.id = s.store_id
                        LEFT JOIN products p ON p.id = s.product_id
                        LEFT JOIN category_classification cc ON cc.third_subcategory_id = p.subcategory_id_3
                    WHERE
                        s.date BETWEEN :fromDate AND :toDate
                        AND st.active = 1
                    GROUP BY
                        s.date,
                        st.name,
                        cc.category_name
                ) AS grouped
                GROUP BY
                    store_name,
                    category_name;
                """;
    }
}
