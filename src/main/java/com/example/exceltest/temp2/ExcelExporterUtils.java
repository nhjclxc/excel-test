package com.example.exceltest.temp2;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class ExcelExporterUtils<T> {
    protected static final Logger log = LoggerFactory.getLogger(ExcelExporterUtils.class);

    public static void main(String[] args) throws IOException {
        String source = "dynamic-template.xlsx";
        // 从resources下加载模板并替换
        InputStream resourceAsStream = ExcelExporterUtils.class.getClassLoader().getResourceAsStream(source);
        // 输出文件路径
        String outputPath = "output -" + source;

        Map<String, Object> testMap = new HashMap<>();
        // 单值格式
        testMap.put("title", "" + "单值填充 -- 这是一个标题噢噢噢");
        testMap.put("signName", "单值填充 -- 张一三");
        testMap.put("age", "单值填充 -- " + 18);
        testMap.put("time", "单值填充 -- " + LocalDateTime.now());
        testMap.put("word", "单值填充 -- 一个单词");
        testMap.put("amount", "单值填充 -- " + new BigDecimal("99999.88"));
        testMap.put("flag", "单值填充 -- " + true);
        testMap.put("hhh", "哈哈哈 -- ");

        // 列表填充格式
        List<TestObj> testObjList = new ArrayList<>();
        testObjList.add(TestObj.builder().id(1).name("张三 ").money(new BigDecimal("111.88")).build());
        testObjList.add(TestObj.builder().id(2).name("里斯 ").money(new BigDecimal("222.66")).build());
        testObjList.add(TestObj.builder().id(3).name("王五 ").money(new BigDecimal("333.88")).build());
        testObjList.add(TestObj.builder().id(4).name("赵六 ").money(new BigDecimal("555.88")).build());
        testObjList.add(TestObj.builder().id(5).name("钱七 ").money(new BigDecimal("666.88")).build());
        testObjList.add(TestObj.builder().id(6).name("测试空值情况 ").build());
        testObjList.add(TestObj.builder().name("测试空值情况 7").build());

        testMap.put(ExcelExporterUtils.LIST_FLAG, testObjList);
        testMap.put(ExcelExporterUtils.CLAZZ_FLAG, TestObj.class);

        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
        exportByTemplate(resourceAsStream, fileOutputStream, testMap);
        fileOutputStream.close();
    }

    public final static String LIST_FLAG = ".";
    public final static String CLAZZ_FLAG = ".CLAZZ_FLAG";


    public static void exportByTemplate(InputStream templateStream, OutputStream outputStream, Map<String, ?> dataMap) throws IOException {

        // 读取模板文件
        assert templateStream != null;
        Workbook workbook = new XSSFWorkbook(templateStream);
//        workbook.sheetIterator().forEachRemaining(sheet -> {
//            // 遍历每一个sheet进行模板数据填充
//            fillCell(sheet, dataMap);
//        });
        Sheet sheet = workbook.getSheetAt(0); // 假设数据在第一个sheet中
        fillCell(sheet, dataMap);

        // 保存到流里面
        workbook.write(outputStream);
        outputStream.flush();
        workbook.close();
    }

    public static void fillCell(Sheet sheet, Map<String, ?> dataMap) {

        // 遍历整个sheet，每一行
        Iterator<Row> rowIterator = sheet.rowIterator();
        int lastRowNum = sheet.getLastRowNum();

        // 将迭代器里面的数据收集到List里面， 便于后续的递归使用 （将Iterator转换为Stream）
        List<Row> allRowList = StreamSupport.stream(Spliterators.spliteratorUnknownSize(rowIterator, Spliterator.ORDERED), false).collect(Collectors.toList());


        int listValueStartRowIndex = -1;
        int dataListSize = 0;

        for (int rowIndex = 0; rowIndex < allRowList.size(); rowIndex++) {
            Row row = allRowList.get(rowIndex);

            // 用于保存是不是有创建新的一行
            List<Row> listRow = new ArrayList<>();

            // 遍历每一个单元格，看看是不是要进行模板数据填充
            Iterator<Cell> cellIterator = row.cellIterator();
            List<Cell> allCellList = StreamSupport.stream(Spliterators.spliteratorUnknownSize(cellIterator, Spliterator.ORDERED), false).collect(Collectors.toList());
            for (Cell cell : allCellList) {
                int columnIndex = cell.getColumnIndex();

                String cellValue = getCellValue(cell);
                // 首先检查是不是单数据填充
                if (isSingleValueFill(cellValue)) {
                    String singleValueFlag = trimSingleValueFillBraces(cellValue);
                    Object setValue = dataMap.get(singleValueFlag);
                    if (setValue != null){
                        cell.setCellValue(setValue.toString());
                    }
                } else if (isListValueFill(cellValue)) {
                    // 接着检查是不是列表元素填充
                    String attribute = trimSingleValueFillBraces(cellValue); // .id
                    String listAttribute = attribute.split("\\.")[1];  //  id
                    List<?> dataList = (List<?>) dataMap.get(ExcelExporterUtils.LIST_FLAG); // 获取list数据
                    Class<?> clazz = (Class<?>) dataMap.get(ExcelExporterUtils.CLAZZ_FLAG); // 获取list对应的泛型，后续使用反射获取值
                    if (dataList == null) {
                        continue;
                    }
                    if (listValueStartRowIndex == -1)
                        listValueStartRowIndex = rowIndex;

                    for (int i = 0; i < dataList.size(); i++) {
                        String fieldValue = getFieldValue(clazz, listAttribute, dataList.get(i));

                        // 创建新的n行，其中n等于列表长度
                        Row newRow;
                        if (listRow.size() < dataList.size()) {
                            // i == 0 直接使用模板这一行，否则在最后创建新的行
                            if (i == 0) {
                                newRow = row;
                                dataListSize = dataList.size();
                            } else {
                                newRow = sheet.createRow(i + lastRowNum);
                            }
                            listRow.add(newRow);
                        } else {
                            newRow = listRow.get(i);
                        }
                        // 将原有单元格样式拿出来保存到副本里面，创建新的单元格之后再设置回去
                        CellStyle cellStyle = cell.getCellStyle();

                        // 创建单元格
                        Cell listRowCell = newRow.createCell(columnIndex);
                        // 从原有单元格中克隆样式 newStyle.cloneStyleFrom(sourceStyle);
                        listRowCell.setCellStyle(cellStyle);
                        // 从列表里面获取对应的值进行填充
                        listRowCell.setCellValue(fieldValue);
                    }
                }
            }
        }

        if (listValueStartRowIndex != -1) {
            // sheet.shiftRows(6, 12, -5);  // 从第6（在excel中指第7行）行到第12（在excel中指第13行）行全部向上移5行

            // 移动单元格
            int moveStart = listValueStartRowIndex + 1; // listValueStartRowIndex表示当前模板行，+1表示从模板行的下一行开始
            int moveEnd = allRowList.size() - 1; // allRowList.size()表示原有模板excel的行数，-1是因为索引，
            int moveSize = (dataListSize - 1) + (moveEnd - moveStart + 1); // (dataList.size() - 1)表示追加到行尾的数据列，(moveEnd - moveStart + 1)表示从原始开头到原始结尾的行数

            // 把原始list填充后面的数据先往最后移动
            sheet.shiftRows(moveStart, moveEnd, moveSize); // moveStart和moveEnd是前闭后闭，

            // 再把list填充（去除第一行）和原始list填充后面的数据整体往前移动，sheet.getLastRowNum()表示单元格的最后一行
            sheet.shiftRows(moveEnd + 1, sheet.getLastRowNum(), -(moveEnd - moveStart + 1));
        }

    }

    private static String getCellValue(Cell cell) {
        String cellValue = "";
        try {
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case STRING:
                    cellValue = String.valueOf(cell.getStringCellValue());
                    break;
                case FORMULA:
                    break;
                case BLANK:
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
            }
        } catch (Exception ignored) {
        }
        return cellValue;
    }

    /**
     * 获取对象对应属性的值
     *
     * @param clazz     泛型
     * @param attribute 属性名称
     * @param obj       对象
     */
    private static String getFieldValue(Class<?> clazz, String attribute, Object obj) {
        try {
            Field field = clazz.getDeclaredField(attribute);
            field.setAccessible(true);
            return field.get(obj).toString();
        } catch (Exception e) {
            log.info(e.getMessage());
        }
        return "";
    }

    /**
     * 判断这个单元格是不是要填充
     * 要填充的单元格必须以"{{"打头和"}}"结尾
     *
     * @param str 单元格的str值
     * @return 如果满足"{{"打头和"}}"结尾则返回true，否则返回false
     * @author 罗贤超
     */
    public static boolean isSingleValueFill(String str) {
        if (str == null || "".equals(str)) {
            return false;
        }
        // (str.startsWith("{{") && str.endsWith("}}"))表示是以"{{"打头"}}"结尾
        // !isListValueFill(str) 表示不是列表填充，注意取了反
        return (str.startsWith("{{") && str.endsWith("}}")) && !isListValueFill(str);
    }

    /**
     * 去除模板里面的"{{"和"}}"，注意调用此方法前，必须调用isSingleValueFill返回为true，才调用这个方法
     *
     * @param str 模板里面的只发出
     * @return 返回的字符串
     * @author 罗贤超
     */
    public static String trimSingleValueFillBraces(String str) {
        return trimDoubleBraces(str, 2, 2);
    }

    public static String trimListValueFillBraces(String str) {
        return trimDoubleBraces(str, 3, 2);
    }

    public static String trimDoubleBraces(String str, int forword, int back) {
        if (str == null || "".equals(str)) {
            return str;
        }
        return str.substring(forword, str.length() - back);
    }

    public static boolean isListValueFill(String str) {
        if (str == null || "".equals(str)) {
            return false;
        }
        return str.startsWith("{{.") && str.endsWith("}}");
    }

}
