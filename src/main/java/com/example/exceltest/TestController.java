package com.example.exceltest;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.util.*;

/**
 * @author LuoXianchao
 * @since 2024/04/21 09:29
 */
@RestController
@RequestMapping("/excel")
public class TestController {


    @GetMapping("test1")
    public void test1(HttpServletResponse response) throws Exception {

        // 单值数据准备
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put("title", "" + "这是一个标题噢噢噢");
        dataMap.put("signName", "张一三");
        dataMap.put("time", "" + LocalDateTime.now());
        dataMap.put("word", "一个单词");
        dataMap.put("who", "不知道谁");
        dataMap.put("abc", "这个abc");

        // 列表数据准备
        List<TestObj> dataList = new ArrayList<>(); // 假设这是你的数据列表
        dataList.add(TestObj.builder().id(1).name("张三 ").age(19).money(new BigDecimal("111.88")).build());
        dataList.add(TestObj.builder().id(2).name("里斯 ").age(29).money(new BigDecimal("222.66")).build());
        dataList.add(TestObj.builder().id(3).name("王五 ").age(39).money(new BigDecimal("333.88")).build());
        dataList.add(TestObj.builder().id(4).name("赵六 ").age(50).money(new BigDecimal("555.88")).build());
        dataList.add(TestObj.builder().id(5).name("钱七 ").age(61).money(new BigDecimal("666.88")).build());

        // 在rescore下面获取文件
        InputStream templateStream = ExcelExporterUtils.class.getClassLoader().getResourceAsStream("dynamic-template.xlsx");
        setResponse(response, "123456输出excel.xlsx");
        // 输出文件
//        ExcelExporterUtils.exportByTemplate(templateStream, response.getOutputStream(), dataMap, dataList, TestObj.class);
    }

    /**
     *
     */
    private void setResponse(HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.addHeader("Access-Control-Allow-Origin", "*");
        response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8.name()));
        response.setContentType("application/octet-stream; charset=UTF-8");
//        IOUtils.write(data, response.getOutputStream());
    }
    @GetMapping("test2")
    public void test2(HttpServletResponse response) throws Exception {
        String source = "dynamic-template-mult-sheet.xlsx";
        // 从resources下加载模板并替换
        InputStream resourceAsStream = ExcelExporterMultSheetUtils.class.getClassLoader().getResourceAsStream(source);
        // 输出
        setResponse(response, System.currentTimeMillis() + "输出excel.xlsx");

        Map<String, Object> dataMap = new HashMap<>();
        // 单值格式
        dataMap.put("title", "" + "单值填充 -- 这是一个标题噢噢噢");
        dataMap.put("signName", "单值填充 -- 张一三");
        dataMap.put("age", "单值填充 -- " + 18);
        dataMap.put("time", "单值填充 -- " + LocalDateTime.now());
        dataMap.put("word", "单值填充 -- 一个单词");
        dataMap.put("amount", "单值填充 -- " + new BigDecimal("99999.88"));
        dataMap.put("flag", "单值填充 -- " + true);
        dataMap.put("hhh", "哈哈哈 -- ");

        // 列表填充格式
        List<TestObj> testObjList = new ArrayList<>();
        testObjList.add(TestObj.builder().id(1).name("张三 ").money(new BigDecimal("111.88")).build());
        testObjList.add(TestObj.builder().id(2).name("里斯 ").money(new BigDecimal("222.66")).build());
        testObjList.add(TestObj.builder().id(3).name("王五 ").money(new BigDecimal("333.88")).build());
        testObjList.add(TestObj.builder().id(4).name("赵六 ").money(new BigDecimal("555.88")).build());
        testObjList.add(TestObj.builder().id(5).name("钱七 ").money(new BigDecimal("666.88")).build());
        testObjList.add(TestObj.builder().id(6).name("测试空值情况 ").build());
        testObjList.add(TestObj.builder().name("测试空值情况 7").build());
        dataMap.put(".list66", testObjList);
        dataMap.put(".list66" + ExcelExporterMultSheetUtils.CLAZZ_FLAG, TestObj.class);

        List<TestDistrict> testDistrictList = new ArrayList<>();
        testDistrictList.add(TestDistrict.builder().name("浙江省").level(1).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("宁波市").level(2).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("江北区").level(3).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("庄市大道").level(4).time(LocalDateTime.now()).build());
        dataMap.put(".list88", testDistrictList);
        dataMap.put(".list88" + ExcelExporterMultSheetUtils.CLAZZ_FLAG, TestDistrict.class);


        ExcelExporterMultSheetUtils.exportByTemplate(resourceAsStream, response.getOutputStream(), dataMap);


    }
    @GetMapping("test3")
    public void test3(HttpServletResponse response) throws Exception {

        // 数据
        List<TestDistrict> testDistrictList = new ArrayList<>();
        testDistrictList.add(TestDistrict.builder().name("中国").level(0).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("浙江省").level(1).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("宁波市").level(2).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("江北区").level(3).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("庄市大道").level(4).time(LocalDateTime.now()).build());

        // 属性与列名对应
        // 注意：这里必须使用LinkedHashMap来确保导出的excel的列有序，
        Map<String, String> map = new LinkedHashMap<>();
        map.put("level", "级别");
        map.put("name", "名称");
        map.put("time", "时间");

//        Workbook export = export(testDistrictList, map, TestDistrict.class);
        Workbook export = ExcelExportCommonUtils.export("导出的表格", 2, 0, "导出的标题", testDistrictList, map, TestDistrict.class);


        setResponse(response, System.currentTimeMillis() + "输出excel.xlsx");
        ServletOutputStream outputStream = response.getOutputStream();
        export.write(outputStream);
        export.close();
    }


}
