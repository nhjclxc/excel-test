package com.example.exceltest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

/**
 * 实现根据Excel文件进行模板导出，并且list填充之后还允许继续填充当值模板数据
 *
 * @author 罗贤超
 * @since 2024/04/21 14:23
 */
public class ExcelExporterMultSheetUtils {
    protected static final Logger log = LoggerFactory.getLogger(ExcelExporterMultSheetUtils.class);

    public static void main(String[] args) throws IOException {
        String source = "dynamic-template-mult-sheet.xlsx";
        // 从resources下加载模板并替换
        InputStream resourceAsStream = ExcelExporterMultSheetUtils.class.getClassLoader().getResourceAsStream(source);
        // 输出文件路径
        String outputPath = "output -" + source;

        Map<String, Object> testMap = new HashMap<>();
        // 单值格式
        testMap.put("title", "" + "这是一个标题噢噢噢");
        testMap.put("signName", "张一三");
        testMap.put("age", 18);
        testMap.put("time", LocalDateTime.now());
        testMap.put("word", "一个单词");
        testMap.put("amount", new BigDecimal("99999.88"));
        testMap.put("flag", true);
        testMap.put("hhh", "哈哈哈");

        // 列表填充格式
        List<TestObj> testObjList = new ArrayList<>();
        testObjList.add(TestObj.builder().id(1).name("张三 ").money(new BigDecimal("111.88")).build());
        testObjList.add(TestObj.builder().id(2).name("里斯 ").money(new BigDecimal("222.66")).build());
        testObjList.add(TestObj.builder().id(3).name("王五 ").money(new BigDecimal("333.88")).build());
        testObjList.add(TestObj.builder().id(4).name("赵六 ").money(new BigDecimal("555.88")).build());
        testObjList.add(TestObj.builder().id(5).name("钱七 ").money(new BigDecimal("666.88")).build());
        testObjList.add(TestObj.builder().id(6).name("测试空值情况 ").build());
        testObjList.add(TestObj.builder().name("测试空值情况 2").build());

        testMap.put(".list66", testObjList);
        // 由于后面要使用反射根据list获取内部属性对应的值，且同时无法直接从list推断出对应的泛型类型，因此这里必须要指明list对应的泛型类型，否则无法导出
        testMap.put(".list66" + ExcelExporterMultSheetUtils.CLAZZ_FLAG, TestObj.class);


        List<TestDistrict> testDistrictList = new ArrayList<>();
        testDistrictList.add(TestDistrict.builder().name("浙江省").level(1).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("宁波市").level(2).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("江北区").level(3).time(LocalDateTime.now()).build());
        testDistrictList.add(TestDistrict.builder().name("庄市大道").level(4).time(LocalDateTime.now()).build());
        testMap.put(".list88", testDistrictList);
        testMap.put(".list88" + ExcelExporterMultSheetUtils.CLAZZ_FLAG, TestDistrict.class);

        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
        exportByTemplate(resourceAsStream, fileOutputStream, testMap);
        fileOutputStream.close();

    }

    /** 字节码标识符 */
    public final static String CLAZZ_FLAG = ".CLAZZ_FLAG";


    /**
     * 根据excel模板文件进行导出
     *
     * @param templateStream 模板文件输入流
     * @param outputStream   数据输出流
     * @param dataMap        数据集
     * @author 罗贤超
     */
    public static void exportByTemplate(InputStream templateStream, OutputStream outputStream, Map<String, ?> dataMap) throws IOException {
        // 读取模板文件
        assert templateStream != null;
        Workbook workbook = new XSSFWorkbook(templateStream);
        workbook.sheetIterator().forEachRemaining(sheet -> {
            // 遍历每一个sheet进行模板数据填充
            fillCell(sheet, dataMap);
        });

        // 输出数据
        workbook.write(outputStream);
        outputStream.flush();
        workbook.close();
    }

    public static void exportByTemplate(String templatePath, OutputStream outputStream, Map<String, ?> dataMap) throws IOException {
        InputStream resourceAsStream = ExcelExporterMultSheetUtils.class.getClassLoader().getResourceAsStream(templatePath);
        exportByTemplate(resourceAsStream, outputStream, dataMap);
    }

    public static ByteArrayOutputStream exportByTemplate(String templatePath, Map<String, ?> dataMap) throws IOException {
        InputStream resourceAsStream = ExcelExporterMultSheetUtils.class.getClassLoader().getResourceAsStream(templatePath);
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        exportByTemplate(resourceAsStream, byteArrayOutputStream, dataMap);
        return byteArrayOutputStream;
    }

    public static ByteArrayOutputStream exportByTemplate(InputStream templateStream, Map<String, ?> dataMap) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        exportByTemplate(templateStream, byteArrayOutputStream, dataMap);
        return byteArrayOutputStream;
    }


    /**
     * 执行单元格数据填充填充
     *
     * @param sheet   sheet页
     * @param dataMap 数据map
     * @author 罗贤超
     */
    private static void fillCell(Sheet sheet, Map<String, ?> dataMap) {

        // 遍历整个sheet，每一行
        Iterator<Row> rowIterator = sheet.rowIterator();
        int lastRowNum = sheet.getLastRowNum();

        int listValueStartRowIndex = -1;
        int dataListSize = 0;

        // 将迭代器里面的数据收集到List里面， 便于后续的递归使用 （将Iterator转换为Stream）
        List<Row> allRowList = StreamSupport.stream(Spliterators.spliteratorUnknownSize(rowIterator, Spliterator.ORDERED), false).collect(Collectors.toList());
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
                    cell.setCellValue(setValue == null ? "" : setValue.toString());
                } else if (isListValueFill(cellValue)) {
                    // 接着检查是不是列表元素填充
                    String attribute = trimSingleValueFillBraces(cellValue); // .list1.id

                    int lastDotIndex = attribute.lastIndexOf('.');
                    String listAttribute = attribute.substring(0, lastDotIndex);  // .list1
                    String listInnerAttribute = attribute.substring(lastDotIndex + 1);  //  id
                    List<?> dataList = (List<?>) dataMap.get(listAttribute); // 获取list数据
                    Class<?> clazz = (Class<?>) dataMap.get(listAttribute + ExcelExporterMultSheetUtils.CLAZZ_FLAG); // 获取list对应的泛型，后续使用反射获取值
                    if (dataList == null) {
                        continue;
                    }
                    if (listValueStartRowIndex == -1)
                        listValueStartRowIndex = rowIndex;

                    for (int i = 0; i < dataList.size(); i++) {
                        String fieldValue = getFieldValue(clazz, listInnerAttribute, dataList.get(i));

                        // 创建新的n行，其中n等于列表长度
                        Row newRow;
                        if (listRow.size() < dataList.size()) {
                            // i == 0 直接使用模板这一行，否则在最后创建新的行
                            if (i == 0) {
                                newRow = row;
                                dataListSize = dataList.size(); // 记录list数据的长度，便于后续单元格的移动
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

        // 移动单元格
        if (listValueStartRowIndex != -1) {
            // sheet.shiftRows(6, 12, -5);  // 从第6（在excel中指第7行）行到第12（在excel中指第13行）行全部向上移5行

            int moveStart = listValueStartRowIndex + 1; // listValueStartRowIndex表示当前模板行，+1表示从模板行的下一行开始
            int moveEnd = allRowList.size() - 1; // allRowList.size()表示原有模板excel的行数，-1是因为索引从0开始
            int moveSize = (dataListSize - 1) + (moveEnd - moveStart + 1); // (dataList.size() - 1)表示追加到行尾的数据列，(moveEnd - moveStart + 1)表示从原始开头到原始结尾的行数

            // 把原始list填充后面的数据先往最后移动
            sheet.shiftRows(moveStart, moveEnd, moveSize); // moveStart和moveEnd都是闭区间，

            // 再把list填充（去除第一行）和原始list填充后面的数据整体往前移动（去除空白单元格），sheet.getLastRowNum()表示单元格的最后一行
            sheet.shiftRows(moveEnd + 1, sheet.getLastRowNum(), -(moveEnd - moveStart + 1));
        }

    }

    /**
     * 获取单元格的值
     *
     * @param cell 单元格
     * @return 单元格原始数据数据
     * @author 罗贤超
     */
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
     * @author 罗贤超
     */
    private static String getFieldValue(Class<?> clazz, String attribute, Object obj) {
        try {
            Field field = clazz.getDeclaredField(attribute);
            field.setAccessible(true);
            return field.get(obj).toString();
        } catch (Exception e) {
            log.info("获取数据为空：{}", e.getMessage());
        }
        return "";
    }

    /**
     * 去除模板里面的"{{"和"}}"，注意调用此方法前，必须调用isSingleValueFill返回为true，才调用这个方法
     *
     * @param str 模板里面的只发出
     * @return 返回的字符串
     * @author 罗贤超
     */
    private static String trimSingleValueFillBraces(String str) {
        if (str == null || "".equals(str)) {
            return str;
        }
        return str.substring(2, str.length() - 2);
    }

    /**
     * 判断这个单元格是不是要填充
     * 要填充的单元格必须以"{{"打头和"}}"结尾
     *
     * @param str 单元格的str值
     * @return 如果满足"{{"打头和"}}"结尾则返回true，否则返回false
     * @author 罗贤超
     */
    private static boolean isSingleValueFill(String str) {
        if (str == null || "".equals(str)) {
            return false;
        }
        // (str.startsWith("{{") && str.endsWith("}}"))表示是以"{{"打头"}}"结尾
        // !isListValueFill(str) 表示不是列表填充，注意取了反
        return (str.startsWith("{{") && str.endsWith("}}")) && !isListValueFill(str);
    }

    private static boolean isListValueFill(String str) {
        if (str == null || "".equals(str)) {
            return false;
        }
        return str.startsWith("{{.") && str.endsWith("}}");
    }

}
