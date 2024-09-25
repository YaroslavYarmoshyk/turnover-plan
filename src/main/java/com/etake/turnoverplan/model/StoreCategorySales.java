package com.etake.turnoverplan.model;


import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;
import lombok.experimental.SuperBuilder;

@Data
@EqualsAndHashCode(callSuper = true)
@SuperBuilder
@ToString(callSuper = true)
public class StoreCategorySales extends StoreSales {
    private String similarStore;
    private String categoryName;
}
