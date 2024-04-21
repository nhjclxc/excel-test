//package com.example.exceltest;
//
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
//
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.io.InputStream;
//import java.lang.reflect.Field;
//import java.math.BigDecimal;
//import java.time.LocalDateTime;
//import java.util.*;
//
//public class ExcelExporterTemp {
//    protected static final Logger log = LoggerFactory.getLogger(ExcelExporterTemp.class);
//
//    public static void main(String[] args) throws IOException, NoSuchFieldException, IllegalAccessException {
//        // 1.从resources下加载模板并替换
////        InputStream resourceAsStream = ExcelExporter.class.getClassLoader().getResourceAsStream("simple-template.xlsx");
//        InputStream resourceAsStream = ExcelExporterTemp.class.getClassLoader().getResourceAsStream("dynamic-template.xlsx");
//
//        // 示例用法
//        List<TestObj> dataList = new ArrayList<>(); // 假设这是你的数据列表
//        dataList.add(TestObj.builder().id(1).name("张三 ").age(18).money(new BigDecimal("66.88")).build());
//        dataList.add(TestObj.builder().id(2).name("里斯 ").age(28).money(new BigDecimal("666.66")).build());
//        dataList.add(TestObj.builder().id(3).name("王五 ").age(38).money(new BigDecimal("888.88")).build());
//
//
//        Map<String, String> map = new HashMap<>();
//        map.put("title", "" + "这是一个标题噢噢噢");
//        map.put("signName", "张一三");
//        map.put("time", "" + LocalDateTime.now());
//        String outputPath = "output.xlsx"; // 输出文件路径
//
//
//        // 读取模板文件
//        assert resourceAsStream != null;
//        Workbook workbook = new XSSFWorkbook(resourceAsStream);
//        Sheet sheet = workbook.getSheetAt(0); // 假设数据在第一个sheet中
//        // 遍历整个sheet，每一行
//        int rowIndex = 0;
//        Iterator<Row> rowIterator = sheet.rowIterator();
//        int lastRowNum = sheet.getLastRowNum();
//        Row removeRow = null;
//        int listValueStartRowIndex = -1;
//        int rowCount = 0;
//        int tempCellIndex = 0;
//        List<Row> saveRow = new ArrayList<>();
//        while (rowIterator.hasNext()) {
//            Row row = rowIterator.next();
//
//            // 用于保存是不是有创建新的一行
//            List<Row> listRow = new ArrayList<>();
//            boolean listValueFlag = false;
//
//            // 遍历每一个单元格，看看是不是要进行模板数据填充
//            int cellIndex = 0;
//            Iterator<Cell> cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//
//                // 首先检查是不是单数据填充
//                String cellValue = cell.getStringCellValue();
//                boolean singleValueFill = isSingleValueFill(cellValue);
//                if (singleValueFill){
//                    String singleValueFlag = trimSingleValueFillBraces(cellValue);
//                    String setValue = map.get(singleValueFlag);
//                    cell.setCellValue(setValue);
//                }
//
//                // 接着检查是不是列表元素填充
//                // 如果是列表填充的话，这一行以后的单元格必须向后移动列表的长度
//                boolean listValueFill = isListValueFill(cellValue);
//                if (listValueFill){
//                    // 标记这一行为列表填充，后续将这一行删除
//                    if (removeRow == null)
//                        removeRow = row;
//                    listValueFlag = true;
//                    listValueStartRowIndex = rowIndex;
//
//                    // 将剩余的模板行保存起来后续遍历
//                    while (rowIterator.hasNext()) {
//                        Row next = rowIterator.next();
//                        saveRow.add(next);
//                    }
//
//                    String attribute = trimListValueFillBraces(cellValue);
//                    for (int i = 0; i < dataList.size(); i++) {
//                        TestObj testObj = dataList.get(i);
//                        String fieldValue = getFieldValue(TestObj.class, attribute, testObj);
//
//                        // 创建新的n行，其中n等于列表长度
//                        Row newRow = null;
//                        if (listRow.size() < dataList.size()) {
//                            // i == 0 直接使用模板这一行，否则创建新的行
//                            newRow = i == 0 ? row : sheet.createRow(i + lastRowNum);
//                            listRow.add(newRow);
//                        }else {
//                             newRow = listRow.get(i);
//                        }
//
//                        // 创建单元格
//                        Cell listRowCell = newRow.createCell(cellIndex);
//                        // 从列表里面获取对应的值进行填充
//                        listRowCell.setCellValue(fieldValue);
//                    }
//
//                }
//
//                cellIndex++;
//            }
//
//            rowCount++;
//            rowIndex++;
//
//            if (listValueFlag){
//                break;
//            }
//        }
//
//        for (Row row : saveRow) {
//            Iterator<Cell> cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//
//                // 首先检查是不是单数据填充
//                String cellValue = cell.getStringCellValue();
//                boolean singleValueFill = isSingleValueFill(cellValue);
//                if (singleValueFill) {
//                    String singleValueFlag = trimSingleValueFillBraces(cellValue);
//                    String setValue = map.get(singleValueFlag);
//                    cell.setCellValue(setValue);
//                }
//            }
//        }
//        System.out.println("rowCount = " + rowCount); //遍历了多少行 4
//        System.out.println("listValueStartRowIndex = " + listValueStartRowIndex); // list填充开始于哪一行 3
//        // 从第6（在excel中指第7行）行到第12（在excel中指第13行）行全部向上移5行
//        // sheet.shiftRows(6, 12, -5);
//        // 把原始list填充后面的数据先往最后移动
//        sheet.shiftRows(listValueStartRowIndex+1, listValueStartRowIndex+saveRow.size(), dataList.size()+1);
//        // 再把list填充（去除第一行）和原始list填充后面的数据整体往前移动
//        sheet.shiftRows(listValueStartRowIndex+1+saveRow.size(), listValueStartRowIndex+1+saveRow.size()+dataList.size(), -(dataList.size()-1));
//
//
//
//
//        // 保存文件
//        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
//        workbook.write(fileOutputStream);
//        workbook.close();
//        fileOutputStream.close();
//
//
//    }
//
//
//    /**
//     * 往对象的对应属性上设值
//     *
//     * @param clazz     泛型
//     * @param attribute 属性名称
//     * @param obj       对象
//     * @param value     值
//     * @param <T>
//     */
//    private static <T> void setFieldValue(Class<T> clazz, String attribute, T obj, Object value) throws NoSuchFieldException, IllegalAccessException {
//        Field field = clazz.getDeclaredField(attribute);
//        field.setAccessible(true);
//        field.set(obj, value);
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
//    private static <T> String getFieldValue(Class<T> clazz, String attribute, T obj){
//        try {
//            Field field = clazz.getDeclaredField(attribute);
//            field.setAccessible(true);
//            return field.get(obj).toString();
//        } catch (NoSuchFieldException | IllegalAccessException e) {
//            log.info(e.getMessage());
//        }
//        return null;
//    }
//
//    /**
//     * 判断这个单元格是不是要填充
//     *      要填充的单元格必须以"{{"打头和"}}"结尾
//     *
//     * @param str 单元格的str值
//     * @return 如果满足"{{"打头和"}}"结尾则返回true，否则返回false
//     * @author 罗贤超
//     */
//    public static boolean isSingleValueFill(String str) {
//        if (str == null || "".equals(str)){
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
//    public static String trimDoubleBraces(String str, int forword, int back) {
//        if (str == null || "".equals(str)){
//            return str;
//        }
//        return str.substring(forword, str.length() - back);
//    }
//
//    public static boolean isListValueFill(String str) {
//        if (str == null || "".equals(str)){
//            return false;
//        }
//        return str.startsWith("{{.") && str.endsWith("}}");
//    }
//
//}
