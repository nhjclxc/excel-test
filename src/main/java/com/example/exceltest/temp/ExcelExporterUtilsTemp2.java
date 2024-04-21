//package com.example.exceltest;
//
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
//
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.io.InputStream;
//import java.io.OutputStream;
//import java.lang.reflect.Field;
//import java.math.BigDecimal;
//import java.time.LocalDateTime;
//import java.util.*;
//import java.util.stream.Collectors;
//import java.util.stream.StreamSupport;
//
///**
// * 实现单值和一个列表的填充
// * @param <T>
// */
//public class ExcelExporterUtilsTemp2<T> {
//    protected static final Logger log = LoggerFactory.getLogger(ExcelExporterUtilsTemp2.class);
//
//    public static void main(String[] args) throws IOException, NoSuchFieldException, IllegalAccessException {
//        // 从resources下加载模板并替换
////        InputStream resourceAsStream = ExcelExporter.class.getClassLoader().getResourceAsStream("simple-template.xlsx");
//        InputStream resourceAsStream = ExcelExporterUtilsTemp2.class.getClassLoader().getResourceAsStream("dynamic-template.xlsx");
//        // 输出文件路径
//        String outputPath = "output.xlsx";
//
//        // 单值数据准备
//        Map<String, String> dataMap = new HashMap<>();
//        dataMap.put("title", "" + "这是一个标题噢噢噢");
//        dataMap.put("signName", "张一三");
//        dataMap.put("time", "" + LocalDateTime.now());
//        dataMap.put("word", "一个单词");
//        dataMap.put("who", "不知道谁");
//        dataMap.put("abc", "这个abc");
//
//        // 列表数据准备
//        List<TestObj> dataList = new ArrayList<>(); // 假设这是你的数据列表
//        dataList.add(TestObj.builder().id(1).name("张三 ").age(19).money(new BigDecimal("111.88")).build());
//        dataList.add(TestObj.builder().id(2).name("里斯 ").age(29).money(new BigDecimal("222.66")).build());
//        dataList.add(TestObj.builder().id(3).name("王五 ").age(39).money(new BigDecimal("333.88")).build());
//        dataList.add(TestObj.builder().id(4).name("赵六 ").age(50).money(new BigDecimal("555.88")).build());
//        dataList.add(TestObj.builder().id(5).name("钱七 ").age(61).money(new BigDecimal("666.88")).build());
////        for (int i = 0; i < 10; i++) {
////            dataList.add(TestObj.builder().id(i*2+10).name("钱七 " + i).age(i*5).money(new BigDecimal("666.88" + i*99)).build());
////        }
//        dataList.add(TestObj.builder().id(6).name("测试空值情况 ").build());
//        dataList.add(TestObj.builder().name("测试空值情况 7").build());
//
//        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
//        exportByTemplate(resourceAsStream, fileOutputStream, dataMap, dataList, TestObj.class);
//        fileOutputStream.close();
//
//
//        Map<String, Object> testMap = new HashMap<>();
//        // 单值格式
//        testMap.put("name", "张三");
//        testMap.put("age", 18);
//        testMap.put("amount", new BigDecimal("99999.88"));
//        testMap.put("flag", true);
//
//        // 多列表填充格式
//        List<TestObj> testObjList = new ArrayList<>();
//        for (int i = 0; i < 10; i++) {
//            testObjList.add(TestObj.builder().id(i * 2 + 10).name("名字啊 " + i).age(i * 5).money(new BigDecimal("666.88" + i * 99)).build());
//        }
//        testMap.put(".testObjList", testObjList);
//        testMap.put(".testObjList" + ExcelExporterUtilsTemp2.CLAZZ_FLAG, TestObj.class);
//
////        List<TestDistrict> testDistrictList = new ArrayList<>();
////        testDistrictList.add(TestDistrict.builder().name("浙江省").level(1).time(LocalDateTime.now()).build());
////        testDistrictList.add(TestDistrict.builder().name("宁波市").level(2).time(LocalDateTime.now()).build());
////        testDistrictList.add(TestDistrict.builder().name("江北区").level(3).time(LocalDateTime.now()).build());
////        testDistrictList.add(TestDistrict.builder().name("庄市大道").level(4).time(LocalDateTime.now()).build());
////        testMap.put(".testDistrictList", testDistrictList);
////        testMap.put(".testObjList" + ExcelExporterUtilsTemp2.CLAZZ_FLAG, TestDistrict.class);
////
//
//    }
//
//    public final static String CLAZZ_FLAG = ".CLAZZ_FLAG";
//
//
//
//
//    public static <T> void exportByTemplate(InputStream templateStream, OutputStream outputStream, Map<String, ?> dataMap, List<T> dataList, Class<T> clazz) throws IOException {
//
//        // 读取模板文件
//        assert templateStream != null;
//        Workbook workbook = new XSSFWorkbook(templateStream);
//        Sheet sheet = workbook.getSheetAt(0); // 假设数据在第一个sheet中
//
//        // 执行模板填充
//        fillCell(sheet, dataMap, dataList, clazz);
//
//        // 保存到流里面
//        workbook.write(outputStream);
//        outputStream.flush();
//        workbook.close();
//    }
//
//    public static <T> void fillCell(Sheet sheet, Map<String, ?> dataMap, List<T> dataList, Class<T> clazz) {
//
//        // 遍历整个sheet，每一行
//        Iterator<Row> rowIterator = sheet.rowIterator();
//        int lastRowNum = sheet.getLastRowNum();
//        int listValueStartRowIndex = -1;
//
//        // 将迭代器里面的数据收集到List里面， 便于后续的递归使用 （将Iterator转换为Stream）
//        List<Row> allRowList = StreamSupport.stream(Spliterators.spliteratorUnknownSize(rowIterator, Spliterator.ORDERED), false).collect(Collectors.toList());
//
//        for (int rowIndex = 0; rowIndex < allRowList.size(); rowIndex++) {
//            Row row = allRowList.get(rowIndex);
//
//            // 用于保存是不是有创建新的一行
//            List<Row> listRow = new ArrayList<>();
//
//            // 遍历每一个单元格，看看是不是要进行模板数据填充
//            int cellIndex = 0;
//            Iterator<Cell> cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//
//                // 首先检查是不是单数据填充
//                String cellValue = cell.getStringCellValue();
//                if (isSingleValueFill(cellValue)) {
//                    String singleValueFlag = trimSingleValueFillBraces(cellValue);
//                    Object setValue = dataMap.get(singleValueFlag);
//                    cell.setCellValue(setValue.toString());
//                } else if (isListValueFill(cellValue)) {
//                    // 接着检查是不是列表元素填充
//                    // 如果是列表填充的话，这一行以后的单元格必须向后移动列表的长度
//                    listValueStartRowIndex = rowIndex;
//
//                    String attribute = trimListValueFillBraces(cellValue);
//                    for (int i = 0; i < dataList.size(); i++) {
//                        String fieldValue = getFieldValue(clazz, attribute, dataList.get(i));
//
//                        // 创建新的n行，其中n等于列表长度
//                        Row newRow;
//                        if (listRow.size() < dataList.size()) {
//                            // i == 0 直接使用模板这一行，否则在最后创建新的行
//                            newRow = i == 0 ? row : sheet.createRow(i + lastRowNum);
//                            listRow.add(newRow);
//                        } else {
//                            newRow = listRow.get(i);
//                        }
//                        // 将原有单元格样式拿出来保存到副本里面，创建新的单元格之后再设置回去
//                        CellStyle cellStyle = cell.getCellStyle();
//
//                        // 创建单元格
//                        Cell listRowCell = newRow.createCell(cellIndex);
//                        // 从原有单元格中克隆样式 newStyle.cloneStyleFrom(sourceStyle);
//                        listRowCell.setCellStyle(cellStyle);
//                        // 从列表里面获取对应的值进行填充
//                        listRowCell.setCellValue(fieldValue);
//                    }
//                }
//                cellIndex++;
//            }
//        }
//
//        if (listValueStartRowIndex != -1) {
//            // 移动单元格
//            // 从第6（在excel中指第7行）行到第12（在excel中指第13行）行全部向上移5行
//            // sheet.shiftRows(6, 12, -5);
//
//            int moveStart = listValueStartRowIndex + 1; // listValueStartRowIndex表示当前模板行，+1表示从模板行的下一行开始
//            int moveEnd = allRowList.size() - 1; // allRowList.size()表示原有模板excel的行数，-1是因为索引，
//            int moveSize = (dataList.size() - 1) + (moveEnd - moveStart + 1); // (dataList.size() - 1)表示追加到行尾的数据列，(moveEnd - moveStart + 1)表示从原始开头到原始结尾的行数
//
//            // 把原始list填充后面的数据先往最后移动
//            sheet.shiftRows(moveStart, moveEnd, moveSize); // moveStart和moveEnd是前闭后闭，
//
//            // 再把list填充（去除第一行）和原始list填充后面的数据整体往前移动，sheet.getLastRowNum()表示单元格的最后一行
//            sheet.shiftRows(moveEnd + 1, sheet.getLastRowNum(), -(moveEnd - moveStart + 1));
//        }
//
//    }
//
//    /**
//     * 获取对象对应属性的值
//     *
//     * @param clazz     泛型
//     * @param attribute 属性名称
//     * @param obj       对象
//     * @param <T>
//     * @return
//     */
//    private static <T> String getFieldValue(Class<T> clazz, String attribute, T obj) {
//        try {
//            Field field = clazz.getDeclaredField(attribute);
//            field.setAccessible(true);
//            return field.get(obj).toString();
//        } catch (Exception e) {
//            log.info(e.getMessage());
//        }
//        return "";
//    }
//
//    /**
//     * 判断这个单元格是不是要填充
//     * 要填充的单元格必须以"{{"打头和"}}"结尾
//     *
//     * @param str 单元格的str值
//     * @return 如果满足"{{"打头和"}}"结尾则返回true，否则返回false
//     * @author 罗贤超
//     */
//    public static boolean isSingleValueFill(String str) {
//        if (str == null || "".equals(str)) {
//            return false;
//        }
//        // (str.startsWith("{{") && str.endsWith("}}"))表示是以"{{"打头"}}"结尾
//        // !isListValueFill(str) 表示不是列表填充，注意取了反
//        return (str.startsWith("{{") && str.endsWith("}}")) && !isListValueFill(str);
//    }
//
//    /**
//     * 去除模板里面的"{{"和"}}"，注意调用此方法前，必须调用isSingleValueFill返回为true，才调用这个方法
//     *
//     * @param str 模板里面的只发出
//     * @return 返回的字符串
//     * @author 罗贤超
//     */
//    public static String trimSingleValueFillBraces(String str) {
//        return trimDoubleBraces(str, 2, 2);
//    }
//
//    public static String trimListValueFillBraces(String str) {
//        return trimDoubleBraces(str, 3, 2);
//    }
//
//    public static String trimDoubleBraces(String str, int forword, int back) {
//        if (str == null || "".equals(str)) {
//            return str;
//        }
//        return str.substring(forword, str.length() - back);
//    }
//
//    public static boolean isListValueFill(String str) {
//        if (str == null || "".equals(str)) {
//            return false;
//        }
//        return str.startsWith("{{.") && str.endsWith("}}");
//    }
//
//}
