package com.etake.turnoverplan.repository.impl;

import com.etake.turnoverplan.model.StoreCategory;
import com.etake.turnoverplan.repository.StoreCategoryRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.jdbc.core.simple.JdbcClient;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
@RequiredArgsConstructor
public class StoreCategoryRepositoryImpl implements StoreCategoryRepository {
    private final JdbcClient jdbcClient;

    @Override
    public List<StoreCategory> getActiveStoreCategoryPairs() {
        return jdbcClient.sql(getActiveStoreCategorySqlQuery())
                .query(StoreCategory.class)
                .list();
    }

    private static String getActiveStoreCategorySqlQuery() {
        return """
                SELECT DISTINCT
                    s.name AS storeName,
                    c.name AS categoryName,
                    r.name AS regionName
                FROM
                    stores s
                        LEFT JOIN
                    regions r ON r.id = s.region_id
                        CROSS JOIN
                    categories c
                WHERE
                    s.active = 1
                  AND c.name != '0-Без категорії'
                  AND c.name != '1-165289'
                  AND r.name != 'Внутрішній';
                """;
    }
}
