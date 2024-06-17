package com.example.exceltest;

import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;


public class ExcelUtils {

    protected static final Logger log = LoggerFactory.getLogger(ExcelUtils.class);

    public static void main(String[] args) throws Exception {
        // 输出文件路径
        String outputPath = "ExcelUtils-test" + ".xlsx";

        // 数据
        List<TestObject> testObjectList = new ArrayList<>();
        testObjectList.add(TestObject.builder().localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(true).build());
        testObjectList.add(TestObject.builder().localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(true).build());
        testObjectList.add(TestObject.builder().localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(false).build());
        testObjectList.add(TestObject.builder().localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(false).build());
        testObjectList.add(TestObject.builder().localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(true).build());

        // 属性与列名对应
        // 注意：这里必须使用LinkedHashMap来确保导出的excel的列有序，
        Map<String, String> map = new LinkedHashMap<>();
        map.put("localDateTime", "localDateTime数据");
        map.put("localDate", "localDate数据");
        map.put("localTime", "localTime数据");
        map.put("date", "date数据");
        map.put("string", "string数据");
        map.put("integer", "integer数据");
        map.put("aFloat", "aFloat数据");
        map.put("aDouble", "aDouble数据");
        map.put("aLong", "aLong数据");
        map.put("bigDecimal", "bigDecimal数据");
        map.put("aBoolean", "aBoolean数据");

        Workbook export = export(testObjectList, map, TestObject.class);
        Workbook workbook = export("导出的表格", 2, 0, "导出的标题", testObjectList, map, TestObject.class);

        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
        workbook.write(fileOutputStream);
        workbook.close();

        List<String> attributeList = new ArrayList<>(map.keySet());
//        List<String> attributeList = Arrays.asList("localDateTime", "localDate", "localTime", "date", "string", "integer", "aFloat", "aDouble", "aLong", "bigDecimal", "aBoolean");
        // 测试导入
        List<TestObject> testObjects = importExcel(new File(outputPath), attributeList, TestObject.class);
        System.out.println(testObjects);
    }



    private ExcelUtils() { }


    /**
     * 检查是不是2003的excel，true表示是2003的excel
     */
    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    /**
     * 检查是不是2007的excel，true表示是2007的excel
     */
    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    /**
     * 判断文件是否合法
     */
    public static boolean validateExcel(String filename) {
        if (filename != null && !"".equals(filename) && (isExcel2003(filename) || isExcel2007(filename))){
            return true;
        }
        throw new RuntimeException("文件名不合法，文件不是[*.xlsx]或[*.xls]");
    }

    public static <T> List<T> importExcel(MultipartFile file, List<String> attributeList, Class<T> clazz) throws IOException {
        return importExcel(file, 0, 2, 0, attributeList, clazz);
    }

    public static <T> List<T> importExcel(File file, List<String> attributeList, Class<T> clazz) throws IOException {
        return importExcel(file, 0, 2, 0, attributeList, clazz);
    }

    public static <T> List<T> importExcel(File file, int sheetIndex, int startRowIndex, int startColumnIndex,
                                          List<String> attributeList, Class<T> clazz) throws IOException {
        String filename = file.getName();
        validateExcel(filename);

        InputStream inputStream = Files.newInputStream(file.toPath());
        boolean excel2003 = isExcel2003(filename);

        return doImportExcel(sheetIndex, startRowIndex, startColumnIndex, inputStream, excel2003, attributeList, clazz);
    }
    public static <T> List<T> importExcel(MultipartFile file, int sheetIndex, int startRowIndex, int startColumnIndex,
                                          List<String> attributeList, Class<T> clazz) throws IOException {
        String filename = file.getOriginalFilename();
        validateExcel(filename);

        InputStream inputStream = file.getInputStream();
        boolean excel2003 = isExcel2003(filename);

        return doImportExcel(sheetIndex, startRowIndex, startColumnIndex, inputStream, excel2003, attributeList, clazz);
    }

    /**
     * 解析excel文件数据
     *
     * @param sheetIndex       要解析的sheet，索引从0开始
     * @param startRowIndex    要解析的起始行，索引从0开始
     * @param startColumnIndex 要解析的起始列，索引从0开始
     * @param inputStream 数据流
     * @param excel2003 是否是2003年版本的excel
     * @param attributeList 对象对应的字段属性名
     * @param clazz 对象对应泛型类型
     * @return 解析得到的数据
     * @author 罗贤超
     */
    public static <T> List<T> doImportExcel(int sheetIndex, int startRowIndex, int startColumnIndex,
                                                           InputStream inputStream, boolean excel2003,
                                                           List<String> attributeList, Class<T> clazz) throws IOException {
        // 创建Workbook
        Workbook workbook = excel2003 ? new HSSFWorkbook(inputStream) : new XSSFWorkbook(inputStream);

        // 获取Sheet
        Sheet sheet = workbook.getSheetAt(sheetIndex);

        List<T> dataList = new ArrayList<>();

//        Iterator<Row> rowIterator = sheet.iterator();
//        for (int rowIndex = 0; rowIterator.hasNext(); rowIndex++) {
//            Row row = rowIterator.next();
        for (Row row : sheet) {
            int rowIndex = row.getRowNum();
            if (rowIndex < startRowIndex) {
                continue;
            }

            Object obj = createObject(clazz);
            for (Cell cell : row) {
                int columnIndex = cell.getColumnIndex();
                if (columnIndex < startColumnIndex) {
                    continue;
                }
                String cellValue = getCellValue(cell);
                String attribute = attributeList.get(columnIndex);
                setFieldValue(clazz, obj, attribute, cellValue);
            }
            // 将obj强制转换为Class<T>类型的对象
            T myObj = clazz.cast(obj);
            dataList.add(myObj);
        }

        workbook.close();
        inputStream.close();

        // 数据预处理 。。。
        return dataList;
    }


    public static <T> Workbook export(List<T> dataList, Map<String, String> attributeMap, Class<T> clazz) {
        return export(null, 1, 0, null, dataList, attributeMap, clazz);
    }

    public static <T> Workbook export(String sheetname, String title, List<T> dataList, Map<String, String> attributeMap, Class<T> clazz) {
        return export(sheetname, 0, 0, title, dataList, attributeMap, clazz);
    }

    public static Workbook export(String sheetname, int freezePaneRow, int freezePaneCol,
                                  String title, List<?> dataList, Map<String, String> attributeMap, Class<?> clazz) {
        sheetname = sheetname == null || "".equals(sheetname) ? "sheet1" : sheetname;
        freezePaneRow = (title != null && !"".equals(title)) ? 2 : 1;
        freezePaneCol = 0;

        return export(-1, sheetname, freezePaneRow, freezePaneCol, title, dataList, attributeMap, clazz);
    }

    public static Workbook export(int rowAccessWindowSize, String sheetname, int freezePaneRow, int freezePaneCol,
                                  String title, List<?> dataList, Map<String, String> attributeMap, Class<?> clazz) {
       /*
        HSSFWorkbook、XSSFWorkbook、SXSSFWorkbook的区别:
         ◎HSSFWorkbook一般用于Excel2003版及更早版本(扩展名为.xls)的导出。上限65535行、256列
         ◎XSSFWorkbook一般用于Excel2007版(扩展名为.xlsx)的导出。上限：1048576行,16384列
         ◎SXSSFWorkbook一般用于大数据量的导出。上限：超出以上两者的限制之后
         */
//      rowAccessWindowSize 显示行上限：-1表示显示所有行，大于0的数据则表示显示设置的函数
        Workbook workbook = new SXSSFWorkbook(rowAccessWindowSize);
        Sheet sheet = workbook.createSheet(sheetname);
        return export(workbook, sheet, freezePaneRow, freezePaneCol, title, dataList, attributeMap, clazz);
    }


    /**
     * excel导出
     *
     * @param workbook      工作簿
     * @param sheet         表格
     * @param freezePaneRow 冻结单元格起始行（索引从0）开始
     * @param freezePaneCol 冻结单元格起始列（索引从0）开始
     * @param dataList      数据集合
     * @param attributeMap   Java对象属性map
     * @param clazz         Java对像的泛型
     * @return .
     * @author 罗贤超
     */
    public static Workbook export(Workbook workbook, Sheet sheet, int freezePaneRow, int freezePaneCol,
                                  String title, List<?> dataList, Map<String, String> attributeMap, Class<?> clazz) {

        // 禁用POI的日志输出
//        Logger.getLogger("org.apache.poi").setLevel(Level.OFF);
//        Logger.getLogger("org.apache.commons").setLevel(Level.OFF);

        sheet.createFreezePane(freezePaneCol, freezePaneRow);

        CellStyle whiteStyle = initCellStyle(workbook, IndexedColors.WHITE);
        CellStyle aquaStyle = initCellStyle(workbook, IndexedColors.AQUA);

        int rowIndex = 0;
        if (title != null && !"".equals(title)) {
            // 添加标题行
            Row titleRow = sheet.createRow(rowIndex++);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue(title);
            titleCell.setCellStyle(whiteStyle);

            // 合并标题单元格
            int size = attributeMap.size();
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, size - 1));
        }

        // 填充表头
        Row tableHeaderRow = sheet.createRow(rowIndex++);
        int headerColumnIndex = 0;
        for (String attribute : attributeMap.keySet()) {
            Cell headerCell = tableHeaderRow.createCell(headerColumnIndex++);
            headerCell.setCellValue(attributeMap.get(attribute));
            headerCell.setCellStyle(whiteStyle);
        }

        // 填充每一行的数据
        for (int i = 0; i < dataList.size(); i++) {
            Row row = sheet.createRow(rowIndex + i);
            Object obj = dataList.get(i);

            // 填充每一个单元格的数据
            int columnIndex = -1;
            for (String attribute : attributeMap.keySet()) {
                // 交替相邻两行的背景颜色
                CellStyle style = (rowIndex + i) % 2 == 0 ? aquaStyle : whiteStyle;

                String fieldValue = getFieldValue(clazz, attribute, obj);
                fillCell(style, row, ++columnIndex, fieldValue);
            }
        }

        // 自适应每一列的单元格大小
        for (int i = 0; i < attributeMap.size() + 1; i++) {
            int width = Math.max(15 * 256, Math.min(255 * 256, sheet.getColumnWidth(i) * 12 / 10));
            sheet.setColumnWidth(i, width);
        }

        return workbook;
    }

    /**
     * 初始化单元格样式
     */
    private static CellStyle initCellStyle(Workbook workbook, IndexedColors colors) {
        CellStyle style = workbook.createCellStyle();
        // 创建边框样式 居中对齐样式等单元格默认样式
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
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(colors.getIndex());
        return style;
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
            // 获取attribute对应的字段
            Field field = getField(clazz, attribute);
            if (field == null){
                throw new NoSuchFieldException("No Such Field Exception !");
            }
            field.setAccessible(true);
            Object o = field.get(obj);
            if (null == o){
                return "";
            }
            Class<?> type = field.getType();
            // 对时间进行格式化
            if (LocalDateTime.class.isAssignableFrom(type)){
                return DateTimeFormatterUtils.LocalDateTimeFormatter.print((LocalDateTime) o, null);
            } else if (LocalDate.class.isAssignableFrom(type)){
                return DateTimeFormatterUtils.LocalDateFormatter.print((LocalDate) o, null);
            } else if (LocalTime.class.isAssignableFrom(type)){
                return DateTimeFormatterUtils.LocalTimeFormatter.print((LocalTime) o, null);
            } else if (Date.class.isAssignableFrom(type)){
                return DateTimeFormatterUtils.DateTimeFormatterCustom.print((Date) o, null);
            } else{
                return o.toString();
            }
        } catch (Exception e) {
            log.info(e.getMessage());
//            System.out.println(e.getMessage());
        }
        return "";
    }

    /**
     * 获取实体类属性，包含继承的属性
     */
    private static <T> Field getField(Class<T> bean, String attribute) throws NoSuchFieldException, IllegalAccessException {
        Class<?> clazz = bean;
        for (; clazz != Object.class; clazz = clazz.getSuperclass()) {//向上循环  遍历父类
            Field[] field = clazz.getDeclaredFields();
            for (Field f : field) {
                f.setAccessible(true);
                if (f.getName().equals(attribute)) {
                    return f;
                }
            }
        }
        return null;
    }


    /**
     * 往对象的对应属性上设值
     *
     * @param clazz     泛型
     * @param obj       对象
     * @param attribute 属性名称
     * @param value     值
     */
    private static <T> void setFieldValue(Class<T> clazz, Object obj, String attribute, Object value) {
        try {
            Field field = clazz.getDeclaredField(attribute); // 获取attribute对应的字段
            field.setAccessible(true); // 允许访问私有属性
//            field.set(obj, value); // 设置属性值

            Class<?> type = field.getType();
            // 对时间进行格式化
            if (LocalDateTime.class.isAssignableFrom(type)){
                value = DateTimeFormatterUtils.LocalDateTimeFormatter.parse((String) value, null);
            } else if (LocalDate.class.isAssignableFrom(type)){
                value = DateTimeFormatterUtils.LocalDateFormatter.parse((String) value, null);
            } else if (LocalTime.class.isAssignableFrom(type)){
                value = DateTimeFormatterUtils.LocalTimeFormatter.parse((String) value, null);
            } else if (Date.class.isAssignableFrom(type)){
                value = DateTimeFormatterUtils.DateTimeFormatterCustom.parse((String) value, null);
            } else if (Integer.class.isAssignableFrom(type)){
                value = Integer.parseInt(value+"");
            } else if (Float.class.isAssignableFrom(type)){
                value = Float.parseFloat(value+"");
            } else if (Double.class.isAssignableFrom(type)){
                value = Double.parseDouble(value+"");
            } else if (Long.class.isAssignableFrom(type)){
                value = Long.parseLong(value+"");
            } else if (Boolean.class.isAssignableFrom(type)){
                value = Boolean.parseBoolean(value+"");
            } else if (BigDecimal.class.isAssignableFrom(type)){
                value = new BigDecimal(value+"");
            }
            field.set(obj, value);
        } catch (NoSuchFieldException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 创建泛型对应的对象
     *  注意：泛型对象必须包含无参构造函数和全参构造函数，可以使用Lombok的两个注解来创建：@NoArgsConstructor和@AllArgsConstructor
     *
     * @param clazz 泛型字节码
     */
    public static <T> T createObject(Class<T> clazz){
        try {
            // 创建对象
            Constructor<T> constructor = clazz.getDeclaredConstructor();
            constructor.setAccessible(true); // 允许访问私有构造函数
            return constructor.newInstance();
        } catch (NoSuchMethodException | InstantiationException | IllegalAccessException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    public static void setResponse(HttpServletResponse response, String fileName) throws IOException {
        response.addHeader("Access-Control-Allow-Origin", "*");
        response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8.name()));
        response.setContentType("application/octet-stream; charset=UTF-8");
    }
}
