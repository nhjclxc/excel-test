package com.example.exceltest.temp2;

import lombok.Builder;
import lombok.Data;

import java.math.BigDecimal;

@Data
@Builder
public class TestObj {
    Integer id;
    String name;
    Integer age;
    BigDecimal money;
}
