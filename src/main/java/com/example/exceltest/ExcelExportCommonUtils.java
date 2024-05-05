package com.example.exceltest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.util.*;

public class ExcelExportCommonUtils<T> {

    protected static final Logger log = LoggerFactory.getLogger(ExcelExportCommonUtils.class);

    public static void main(String[] args) throws IOException {
        // 输出文件路径
        String outputPath = "ExcelExportCommonUtils-test" + ".xlsx";

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
        map.put("time", "时间");
        map.put("name", "名称");
        map.put("level", "级别");

//        Workbook export = export(testDistrictList, map, TestDistrict.class);
        Workbook workbook = export("导出的表格", 2, 0, "导出的标题", testDistrictList, map, TestDistrict.class);

        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
        workbook.write(fileOutputStream);
        workbook.close();
    }


    public static <T> Workbook export(List<T> dataList, Map<String, String> propertyMap, Class<T> clazz) {
        return export(null, 1, 0, null, dataList, propertyMap, clazz);
    }

    public static <T> Workbook export(String sheetname, String title, List<T> dataList, Map<String, String> propertyMap, Class<T> clazz) {
        return export(sheetname, 0, 0, title, dataList, propertyMap, clazz);
    }

    public static Workbook export(String sheetname, int freezePaneRow, int freezePaneCol,
                                  String title, List<?> dataList, Map<String, String> propertyMap, Class<?> clazz){
        sheetname = sheetname == null || "".equals(sheetname) ? "sheet1" : sheetname;
        freezePaneRow = (title != null && !"".equals(title)) ? 2 : 1;
        freezePaneCol = 0;

        return export(-1, sheetname, freezePaneRow, freezePaneCol, title, dataList, propertyMap, clazz);
    }

    public static Workbook export(int rowAccessWindowSize, String sheetname, int freezePaneRow, int freezePaneCol,
                                  String title, List<?> dataList, Map<String, String> propertyMap, Class<?> clazz){
       /*
        HSSFWorkbook、XSSFWorkbook、SXSSFWorkbook的区别:
         ◎HSSFWorkbook一般用于Excel2003版及更早版本(扩展名为.xls)的导出。上限65535行、256列
         ◎XSSFWorkbook一般用于Excel2007版(扩展名为.xlsx)的导出。上限：1048576行,16384列
         ◎SXSSFWorkbook一般用于大数据量的导出。上限：超出以上两者的限制之后
         */
//      rowAccessWindowSize 显示行上限：-1表示显示所有行，大于0的数据则表示显示设置的函数
        Workbook workbook = new SXSSFWorkbook(rowAccessWindowSize);
        Sheet sheet = workbook.createSheet(sheetname);
        return export(workbook, sheet, freezePaneRow, freezePaneCol, title, dataList, propertyMap, clazz);
    }


    /**
     * excel导出
     *
     * @param workbook 工作簿
     * @param sheet 表格
     * @param freezePaneRow 冻结单元格起始行（索引从0）开始
     * @param freezePaneCol 冻结单元格起始列（索引从0）开始
     * @param dataList 数据集合
     * @param propertyMap Java对象属性map
     * @param clazz Java对像的泛型
     * @return .
     * @author 罗贤超
     */
    public static Workbook export(Workbook workbook, Sheet sheet, int freezePaneRow, int freezePaneCol,
                                     String title, List<?> dataList, Map<String, String> propertyMap, Class<?> clazz) {

        // 禁用POI的日志输出
//        Logger.getLogger("org.apache.poi").setLevel(Level.OFF);
//        Logger.getLogger("org.apache.commons").setLevel(Level.OFF);

        sheet.createFreezePane(freezePaneCol, freezePaneRow);

        // 创建边框样式 居中对齐样式等单元格默认样式
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(false); // 是否自动换行
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        // 前景颜色
//        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        int rowIndex = 0;
        if (title != null && !"".equals(title)){
            // 添加标题行
            Row titleRow = sheet.createRow(rowIndex++);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue(title);
            titleCell.setCellStyle(style);

            // 合并标题单元格
            int size = propertyMap.size();
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, size - 1));
        }

        // 填充表头
        Row tableHeaderRow = sheet.createRow(rowIndex++);
        int headerColumnIndex = 0;
        for (String property : propertyMap.keySet()) {
            Cell headerCell = tableHeaderRow.createCell(headerColumnIndex++);
            headerCell.setCellValue(propertyMap.get(property));
            headerCell.setCellStyle(style);
        }

        // 填充每一行的数据
        for (int i = 0; i < dataList.size(); i++) {
            Row row = sheet.createRow(rowIndex + i);
            Object obj = dataList.get(i);

            // 填充每一个单元格的数据
            int columnIndex = -1;
            for (String property : propertyMap.keySet()) {
                // 交替相邻两行的背景颜色
//                style.setFillForegroundColor((rowIndex + i) % 2 == 0 ? IndexedColors.AQUA.getIndex() : IndexedColors.YELLOW.getIndex());

                String fieldValue = getFieldValue(clazz, property, obj);
                fillCell(style, row, ++columnIndex, fieldValue);
            }
        }

        // 自适应每一列的单元格大小
        for (int i = 0; i < propertyMap.size() + 1; i++) {
            int width = Math.max(15 * 256, Math.min(255 * 256, sheet.getColumnWidth(i) * 12 / 10));
            sheet.setColumnWidth(i, width);
        }

        return workbook;
    }

    /**
     * 填充单元格数据
     */
    private static void fillCell(CellStyle style, Row row, int columnIndex, String value) {
        Cell cell1 = row.createCell(columnIndex);
        cell1.setCellStyle(style);
        cell1.setCellValue(value);
    }

    /**
     * 获取单元格数据
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
     */
    private static String getFieldValue(Class<?> clazz, String attribute, Object obj) {
        try {
            Field field = clazz.getDeclaredField(attribute);
            field.setAccessible(true);
            return field.get(obj).toString();
        } catch (Exception e) {
            log.info(e.getMessage());
//            System.out.println(e.getMessage());
        }
        return "";
    }

}
