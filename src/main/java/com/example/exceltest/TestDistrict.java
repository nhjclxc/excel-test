package com.example.exceltest;

import lombok.Builder;
import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDateTime;

@Data
@Builder
public class TestDistrict {
    String name;
    Integer level;
    LocalDateTime time;
}
