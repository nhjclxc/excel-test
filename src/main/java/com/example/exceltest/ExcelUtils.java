package com.example.exceltest;

//import cn.hutool.http.HttpRequest;
//import cn.hutool.http.HttpResponse;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.springframework.http.HttpHeaders;
import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.awt.image.MemoryImageSource;
import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;


/**
 * excel表格数据
 * @date 2024-11-30：新增导出Excel最后一列添加图片的功能
 * @author 罗贤超
 */
public class ExcelUtils {

    protected static final Logger log = LoggerFactory.getLogger(ExcelUtils.class);

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
    private static <T> List<T> doImportExcel(int sheetIndex, int startRowIndex, int startColumnIndex,
                                                           InputStream inputStream, boolean excel2003,
                                                           List<String> attributeList, Class<T> clazz) throws IOException {
        // 创建Workbook
        Workbook workbook = excel2003 ? new HSSFWorkbook(inputStream) : new XSSFWorkbook(inputStream);

        // 获取excel里面保存的二级制图片数据（）注意：当前只支持读取图片数据
        List<? extends PictureData> pictureList = workbook.getAllPictures();

        // 获取Sheet
        Sheet sheet = workbook.getSheetAt(sheetIndex);

        List<T> dataList = new ArrayList<>();

//        Iterator<Row> rowIterator = sheet.iterator();
//        for (int rowIndex = 0; rowIterator.hasNext(); rowIndex++) {
//            Row row = rowIterator.next();
        int attributeListSize = attributeList.size();
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
                if (columnIndex >= attributeListSize)
                    break;
                String cellValue = getCellValue(cell);
                String attribute = attributeList.get(columnIndex);

                if (setBinaryData(clazz, obj, attribute, pictureList, rowIndex - startRowIndex)) {
                    // 如果是二级制数据那么文本数据就不要设置文本数据了
                    continue;
                }

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


    /**
     * 设置单元格的二级制数据到对象里面
     * <br>
     * <b>注意：此操作目前仅支持一行对应一个图片数据的情况</b>
     */
    private static <T> boolean setBinaryData(Class<T> clazz, Object obj, String attribute,
                                             List<? extends PictureData> pictureList, int rowIndex) {

        if (null == pictureList || pictureList.size() == 0 || rowIndex >= pictureList.size()) {
            return false;
        }

        try {
            // 获取attribute对应的字段
            Field field = getField(clazz, attribute);
            field.setAccessible(true); // 允许访问私有属性

            Class<?> fieldType = field.getType();

            // 判断当前这个属性是不是图片数据，是图片数据才去获取图片数据
            //  如何变量名称是base64如何读取??? 如imgBase64，是否应该把图片数据也放进去呢？
            if (!(isBinaryData(fieldType) || (attribute != null && attribute.toLowerCase().contains("BASE64".toLowerCase())) ) ) {
                return false;
            }

            // 去读取
            PictureData picture = pictureList.get(rowIndex);
            // 获取图片的字节数据
            byte[] data = picture.getData();
//            String ext = picture.suggestFileExtension();  // 确定图片格式

            // 数据格式转化
            if ( attribute != null && attribute.toLowerCase().contains("BASE64".toLowerCase()) ) {
                // 0、base64字符串，不包含前缀
                String base64 = Base64.getEncoder().encodeToString(data);
                field.set(obj, base64);
            } else if (fieldType.isArray() && fieldType.getComponentType() == byte.class) {
                // 1、byte[]
                field.set(obj, data);
            } else if (fieldType.isArray() && fieldType.getComponentType() == Byte.class) {
                // 2、Byte[]
                // 使用 Stream 将 byte[] 转换为 Byte[]
                Byte[] byteWrapperArray = IntStream.range(0, data.length)
                        .mapToObj(i -> data[i])  // 自动装箱
                        .toArray(Byte[]::new);
                field.set(obj, byteWrapperArray);
            } else if (File.class.isAssignableFrom(fieldType)) {
                // 3、File
                // 获取系统临时目录
                File tempFile = createTempFile();
                try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                    fos.write(data);
                }
                // 可选：自动删除临时文件（在JVM退出时删除）
//                tempFile.deleteOnExit();
                field.set(obj, tempFile);
            } else if (InputStream.class.isAssignableFrom(fieldType) || ByteArrayInputStream.class.isAssignableFrom(fieldType)) {
                // 4、InputStream 5、ByteArrayInputStream
                field.set(obj, new ByteArrayInputStream(data));
            } else if (Image.class.isAssignableFrom(fieldType)) {
                // 6、Image
                Image image = Toolkit.getDefaultToolkit().createImage(data);
                field.set(obj, image);
            } else if (BufferedImage.class.isAssignableFrom(fieldType)) {
                // 6、BufferedImage
                BufferedImage bufferedImage = ImageIO.read(new ByteArrayInputStream(data));
                field.set(obj, bufferedImage);
            } else if (MemoryImageSource.class.isAssignableFrom(fieldType)) {
                // 6、MemoryImageSource

                // 将字节数组转为 BufferedImage
                BufferedImage bufferedImage = ImageIO.read(new ByteArrayInputStream(data));

                if (bufferedImage != null) {
                    // 获取图片的宽高
                    int width = bufferedImage.getWidth();
                    int height = bufferedImage.getHeight();

                    // 这里假设每个像素是4字节（RGBA格式），根据需要调整
                    int[] pixels = new int[width * height];
                    for (int i = 0; i < data.length; i++) {
                        // 简化处理，您需要根据实际的图像格式来解码数据
                        pixels[i] = data[i] & 0xFF;
                    }
                    MemoryImageSource memoryImageSource = new MemoryImageSource(width, height, pixels, 0, width);
                    field.set(obj, memoryImageSource);
                }
            } else if (ByteBuffer.class.isAssignableFrom(fieldType) || FileChannel.class.isAssignableFrom(fieldType)) {
                File tempFile = createTempFile();

                try (FileOutputStream fos = new FileOutputStream(tempFile);
                    FileChannel fileChannel = fos.getChannel()) {
                    ByteBuffer buffer = ByteBuffer.wrap(data);
                    fileChannel.write(buffer);
                    field.set(obj, ByteBuffer.class.isAssignableFrom(fieldType) ? buffer : fileChannel);
                }
                // 转化为内存的FileOutputStream后，临时文件删除
                tempFile.deleteOnExit();
            } else {
                return false;
            }

            return true;
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * 创建临时文件
     */
    private static File createTempFile() throws IOException {
        // 获取系统临时目录
        String tempDir = System.getProperty("java.io.tmpdir");
        // 创建临时文件
        return File.createTempFile("tempFile_" + UUID.randomUUID().toString().replaceAll("-", ""), ".txt", new File(tempDir));
    }


    /**
     * 判断是否为二级制数据
     */
    public static boolean isBinaryData(Class<?> fieldType) {
        return  isArrayOfByte(fieldType)
                || File.class.isAssignableFrom(fieldType) || InputStream.class.isAssignableFrom(fieldType) || ByteArrayInputStream.class.isAssignableFrom(fieldType)
                || Image.class.isAssignableFrom(fieldType) || BufferedImage.class.isAssignableFrom(fieldType) || MemoryImageSource.class.isAssignableFrom(fieldType)
                || ByteBuffer.class.isAssignableFrom(fieldType) || FileChannel.class.isAssignableFrom(fieldType);
    }

    /**
     * 判断是否为字节数组
     */
    public static boolean isArrayOfByte(Class<?> fieldType) {
        // byte[].class.isAssignableFrom(type) || Byte[].class.isAssignableFrom(type)
        return fieldType.isArray() && (fieldType.getComponentType() == byte.class || fieldType.getComponentType() == Byte.class);
    }


    public static <T> Workbook export(List<T> dataList, Map<String, String> attributeMap, Class<T> clazz, boolean rowAlternatelyStyle) {
        return export(null, 1, 0, null, dataList, attributeMap, clazz, rowAlternatelyStyle);
    }

    public static <T> Workbook export(String sheetname, String title, List<T> dataList, Map<String, String> attributeMap, Class<T> clazz, boolean rowAlternatelyStyle) {
        return export(sheetname, 0, 0, title, dataList, attributeMap, clazz, rowAlternatelyStyle);
    }

    public static Workbook export(String sheetname, int freezePaneRow, int freezePaneCol,
                                  String title, List<?> dataList, Map<String, String> attributeMap, Class<?> clazz, boolean rowAlternatelyStyle) {
        sheetname = sheetname == null || "".equals(sheetname) ? "sheet1" : sheetname;
        freezePaneRow = (title != null && !"".equals(title)) ? 2 : 1;
        freezePaneCol = 0;

        return export(-1, sheetname, freezePaneRow, freezePaneCol, title, dataList, attributeMap, clazz, rowAlternatelyStyle);
    }

    public static Workbook export(int rowAccessWindowSize, String sheetname, int freezePaneRow, int freezePaneCol,
                                  String title, List<?> dataList, Map<String, String> attributeMap, Class<?> clazz, boolean rowAlternatelyStyle) {
       /*
        HSSFWorkbook、XSSFWorkbook、SXSSFWorkbook的区别:
         ◎HSSFWorkbook一般用于Excel2003版及更早版本(扩展名为.xls)的导出。上限65535行、256列
         ◎XSSFWorkbook一般用于Excel2007版(扩展名为.xlsx)的导出。上限：1048576行,16384列
         ◎SXSSFWorkbook一般用于大数据量的导出。上限：超出以上两者的限制之后
         */
//      rowAccessWindowSize 显示行上限：-1表示显示所有行，大于0的数据则表示显示设置的函数
        Workbook workbook = new SXSSFWorkbook(rowAccessWindowSize);
        Sheet sheet = workbook.createSheet(sheetname);
        return export(workbook, sheet, freezePaneRow, freezePaneCol, title, dataList, attributeMap, clazz, rowAlternatelyStyle, null, null, -1, -1);
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
    public static Workbook export(Workbook workbook, Sheet sheet, int freezePaneRow, int freezePaneCol, String title,
                                  List<?> dataList, Map<String, String> attributeMap, Class<?> clazz, boolean rowAlternatelyStyle,
                                  String imageUrlAttribute, Map<String, InputStream> inputStreamMap,
                                  int imageHeight, int imageWidth) {

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

        boolean hasImage = null != imageUrlAttribute && !"".equals(imageUrlAttribute) && null != inputStreamMap;

        //生成用于插入图片的容器--这个方法返回的类型在老api中不同
        Drawing<?> drawingPatriarch = null;
        if (hasImage){
            drawingPatriarch = sheet.createDrawingPatriarch();
            // 设置图片一列的列宽
            sheet.setColumnWidth(attributeMap.size() + 1, imageWidth * 50);
        }

        // 填充每一行的数据
        for (int i = 0; i < dataList.size(); i++) {
            int currentRowIndex = rowIndex + i;
            Row row = sheet.createRow(currentRowIndex);
            Object obj = dataList.get(i);

            // 填充每一个单元格的数据
            int columnIndex = -1;
            for (String attribute : attributeMap.keySet()) {
                // 交替相邻两行的背景颜色
                CellStyle style = rowAlternatelyStyle ? (currentRowIndex % 2 == 0 ? aquaStyle : whiteStyle) : whiteStyle;
                String fieldValue = getFieldValue(clazz, attribute, obj);
                fillCell(style, row, ++columnIndex, fieldValue);
            }

            // 最后一列添加图片数据
            if (hasImage) {
                fillCell(whiteStyle, row, ++columnIndex, "");
                // 设置图片行高
                row.setHeightInPoints(imageHeight);

                String imageUrl = ExcelUtils.getFieldValue(clazz, imageUrlAttribute, dataList.get(i));
                if (inputStreamMap.containsKey(imageUrl)) {
                    InputStream inputStream = inputStreamMap.get(imageUrl);
                    createPicture(inputStream, workbook, drawingPatriarch,
                            0, 0, 1023, 255,
                            columnIndex, currentRowIndex,
                            columnIndex + 1, currentRowIndex + 1);
                }
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
     * 在Excel最后一列插入图片
     *
     * @param imageUrlAttribute 图片地址属性
     * @param imageHeight 图片高度
     * @param imageWidth 图片宽度
     * @author 罗贤超
     */
    public static Workbook exportByImage(String sheetname, int freezePaneRow, int freezePaneCol, String title,
                                         List<?> dataList, Map<String, String> attributeMap, Class<?> clazz,
                                         boolean rowAlternatelyStyle, String imageUrlAttribute,
                                         int imageHeight, int imageWidth) {

        // 先异步下载所有图片数据
        Map<String, InputStream> inputStreamMap = new HashMap<>();
        CountDownLatch latch = new CountDownLatch(dataList.size());

        int threadCount = 4; // 设置线程数
        ExecutorService executorService = Executors.newFixedThreadPool(threadCount);

        for (Object o : dataList) {
            String imageUrl = ExcelUtils.getFieldValue(clazz, imageUrlAttribute, o);
            executorService.submit(() -> {
                try {
                    // 判断某个链接指向的数据是否以及存在，存在则不下载，反之则去下载图片数据
                    if (!inputStreamMap.containsKey(imageUrl)) {
                        InputStream inputStream = doDownloadData(imageUrl);
                        if (inputStream != null) {
//                            System.out.println(imageUrl);
                            inputStreamMap.putIfAbsent(imageUrl, inputStream);
                        }
                    }
                } catch (IOException ignored) {
                } finally {
                    latch.countDown(); // 完成一个任务，latch减1
                }
            });
        }
        try { latch.await(); } catch (InterruptedException ignored) { }
        executorService.shutdown();

        // 接着再去写表格
        Workbook workbook = new SXSSFWorkbook(-1);
        Sheet sheet = workbook.createSheet(sheetname);
        export(workbook, sheet, freezePaneRow, freezePaneCol, title, dataList, attributeMap, clazz, rowAlternatelyStyle,
                imageUrlAttribute, inputStreamMap, imageHeight, imageWidth);

        return workbook;
    }



    /**
     * 将inputStream对应的图片数据插入以下指定位置
     *
     * @param inputStream 图片流数据
     * @param workbook 工作簿
     * @param drawingPatriarch 工作簿画图对象
     * @param dx1 起始单元格内的x偏移（0-1023）
     * @param dy1 起始单元格内的y偏移（0-255）
     * @param dx2 终止单元格内的x偏移（0-1023）
     * @param dy2 终止单元格内的y偏移（0-255）
     * @param col1 起始列，表示图片的左上角所在的起始列索引（从 0 开始计数）。
     * @param row1 起始行，表示图片的左上角所在的起始行索引（从 0 开始计数）。
     * @param col2 终止列，表示图片的右下角所在的终止列索引（从 0 开始计数）。
     * @param row2 终止行，表示图片的右下角所在的终止行索引（从 0 开始计数）。
     * @author 罗贤超
     */
    public static void createPicture(InputStream inputStream, Workbook workbook, Drawing<?> drawingPatriarch,
                                     int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2) {
//        row.setHeightInPoints(255); // 设置某行的高度
//        final Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch(); // 获取画图对象
//        XSSFClientAnchor anchor = new XSSFClientAnchor(
//                100, 50,       // dx1, dy1: 起始偏移
//                200, 100,      // dx2, dy2: 终止偏移
//                1, 1,          // col1, row1: 起始列和行（索引从 0 开始）
//                3, 2           // col2, row2: 终止列和行
//        );
        if (null == drawingPatriarch) {
            return;
        }

        try {
            if (inputStream.available() <= 0) {
                return;
            }

            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            ImageIO.write(ImageIO.read(inputStream), "jpg", byteArrayOut);
            //设置每张图片插入位置
            final XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
            // 插入图片
            drawingPatriarch.createPicture(anchor, workbook.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            byteArrayOut.close();
        } catch (IOException ignored) {
        }
    }


    /**
     * 根据urlList下载图片数据
     *
     * @param urlList 所有图片数据链接
     * @author 罗贤超
     */
    public static Map<String, InputStream> downloadData(List<String> urlList) throws InterruptedException {
//        List<String> urlList = Arrays.asList("http://example.com/file1", "http://example.com/file2"); // 示例URL列表
//        Map<String, InputStream> stringInputStreamMap = ExcelUtils.downloadData(urls);

        int threadCount = 4; // 设置线程数

        ExecutorService executorService = Executors.newFixedThreadPool(threadCount);
        CountDownLatch latch = new CountDownLatch(urlList.size());

        Map<String, InputStream> inputStreamMap = new HashMap<>();
        urlList.forEach(url -> executorService.submit(() -> {
            try {
                InputStream inputStream = doDownloadData(url);
                if (inputStream != null) {
                    inputStreamMap.put(url, inputStream);
                }
            } catch (IOException ignored) {
            } finally {
                latch.countDown(); // 完成一个任务，latch减1
            }
        }));

        latch.await(); // 等待所有任务完成
        executorService.shutdown();
        System.out.println("所有文件下载完成");
        return inputStreamMap;
    }

    public static InputStream doDownloadData(String imageUrl) throws IOException {
//            HttpResponse response = HttpRequest.get(imageUrl).execute();
//            if (response.bodyStream().available() > 0) {
//                return response.bodyStream();
//            }
        if (null == imageUrl || "".equals(imageUrl)) {
            return null;
        }
        URL url = new URL(imageUrl);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        connection.connect();
        // 获取输入流
        InputStream inputStream = connection.getInputStream();
//        if (connection.getResponseCode() != HttpURLConnection.HTTP_OK || inputStream.available() <= 0) {
        if (inputStream.available() <= 0) {
            return null;
        }
        return inputStream;
    }

    /**
     * 初始化单元格样式
     */
    public static CellStyle initCellStyle(Workbook workbook, IndexedColors colors) {
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


    /** excel纯数字读取为字符串格式化 */
    private static final DecimalFormat numberFormat = new DecimalFormat("#");

    /**
     * 获取单元格数据
     */
    private static String getCellValue(Cell cell) {
        String cellValue = "";
        try {
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case NUMERIC:
                    cellValue = numberFormat.format(cell.getNumericCellValue());
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
    public static String getFieldValue(Class<?> clazz, String attribute, Object obj) {
        try {

            // 获取attribute对应的字段
            Field field = getField(clazz, attribute);
//            Field field = clazz.getDeclaredField(attribute);
            field.setAccessible(true);
            Object o = field.get(obj);
            if (null == o){
                return "";
            }
            Class<?> type = field.getType();
            // 对时间进行格式化
            if (LocalDateTime.class.isAssignableFrom(type)){
                return DateTimeFormatterUtils.LocalDateTimeFormatter.print((LocalDateTime) o);
            } else if (LocalDate.class.isAssignableFrom(type)){
                return DateTimeFormatterUtils.LocalDateFormatter.print((LocalDate) o);
            } else if (LocalTime.class.isAssignableFrom(type)){
                return DateTimeFormatterUtils.LocalTimeFormatter.print((LocalTime) o);
            } else if (Date.class.isAssignableFrom(type)){
                return DateTimeFormatterUtils.DateTimeFormatterCustom.print((Date) o);
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
     * 往对象的对应属性上设值
     *
     * @param clazz     泛型
     * @param obj       对象
     * @param attribute 属性名称
     * @param value     值
     */
    private static <T> void setFieldValue(Class<T> clazz, Object obj, String attribute, Object value) {
        try {
            if (value == null || "".equals(value)){
                return;
            }
            // 获取attribute对应的字段
            Field field = getField(clazz, attribute);
            field.setAccessible(true); // 允许访问私有属性
//            field.set(obj, value); // 设置属性值

            Class<?> type = field.getType();
            // 对时间进行格式化
            if (LocalDateTime.class.isAssignableFrom(type)){
                value = DateTimeFormatterUtils.LocalDateTimeFormatter.parse((String) value);
            } else if (LocalDate.class.isAssignableFrom(type)){
                value = DateTimeFormatterUtils.LocalDateFormatter.parse((String) value);
            } else if (LocalTime.class.isAssignableFrom(type)){
                value = DateTimeFormatterUtils.LocalTimeFormatter.parse((String) value);
            } else if (Date.class.isAssignableFrom(type)){
                value = DateTimeFormatterUtils.DateTimeFormatterCustom.parse((String) value);
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
            } else if (String.class.isAssignableFrom(type)){
                value = value + "";
            }
            field.set(obj, value);
        } catch (NoSuchFieldException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
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
        throw new NoSuchFieldException("No Such Field Exception !");
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

    public static CellStyle initDefaultCellStyle(Workbook workbook) {
        // 创建边框样式 居中对齐样式
        CellStyle commonStyle = workbook.createCellStyle();
        commonStyle.setBorderBottom(BorderStyle.THIN);
        commonStyle.setBorderTop(BorderStyle.THIN);
        commonStyle.setBorderRight(BorderStyle.THIN);
        commonStyle.setBorderLeft(BorderStyle.THIN);
        commonStyle.setAlignment(HorizontalAlignment.CENTER);
        commonStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return commonStyle;
    }

    public static void setResponse(HttpServletResponse response, String fileName) throws IOException {
//        response.addHeader("Access-Control-Allow-Origin", "*");
//        response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
//        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8.name()));
//        response.setContentType("application/octet-stream; charset=UTF-8");
        response.reset();
        response.addHeader(HttpHeaders.ACCESS_CONTROL_ALLOW_ORIGIN, "*");
        response.addHeader(HttpHeaders.ACCESS_CONTROL_ALLOW_METHODS, "*");
        response.addHeader(HttpHeaders.ACCESS_CONTROL_ALLOW_HEADERS, "*");
        response.addHeader(HttpHeaders.ACCESS_CONTROL_EXPOSE_HEADERS, HttpHeaders.CONTENT_DISPOSITION);
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8.name()));
        response.setContentType("application/octet-stream; charset=UTF-8");
    }


    private static class DateTimeFormatterUtils {
        // Text '45690.71575231481' could not be parsed at index 0
        private final static Pattern LOCAL_DATE_PATTERN = Pattern.compile("Text '([\\d.]+)' could not be parsed");
        // Unparseable date: "45690.71575231481"
        private final static Pattern DATE_PATTERN = Pattern.compile("Unparseable date: '([\\d.]+)'");

        public static class LocalDateTimeFormatter {
            private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            public static LocalDateTime parse(String text) {
                LocalDateTime parse;
                try {
                    parse = LocalDateTime.parse(text, dateTimeFormatter);
                } catch (Exception e) {
                    // 使用正则提取单引号中的数字
                    Matcher matcher = LOCAL_DATE_PATTERN.matcher(e.getMessage());

                    if (!matcher.find()) {
                        throw new RuntimeException(e);
                    }
                    Date date = DateUtil.getJavaDate(Double.parseDouble(text)); // 转为 java.util.Date
                    parse = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
                }
                return parse;
            }
            public static String print(LocalDateTime object) {  return dateTimeFormatter.format(object); }
        }
        public static class LocalDateFormatter {
            private static final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
            public static LocalDate parse(String text) {
                LocalDate parse;
                try {
                    parse = LocalDate.parse(text, dateFormatter);
                } catch (Exception e) {

                    // 使用正则提取单引号中的数字
                    Matcher matcher = LOCAL_DATE_PATTERN.matcher(e.getMessage());

                    if (!matcher.find()) {
                        throw new RuntimeException(e);
                    }
                    Date date = DateUtil.getJavaDate(Double.parseDouble(text)); // 转为 java.util.Date
                    parse = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                }
                return parse;
            }
            public static String print(LocalDate object) {  return dateFormatter.format(object);  }
        }
        public static class LocalTimeFormatter {
            private static final DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");
            public static LocalTime parse(String text) {
                LocalTime parse;
                try {
                    parse = LocalTime.parse(text, timeFormatter);
                } catch (Exception e) {
                    // 使用正则提取单引号中的数字
                    Matcher matcher = LOCAL_DATE_PATTERN.matcher(e.getMessage());
                    if (!matcher.find()) {
                        throw new RuntimeException(e);
                    }
                    Date date = DateUtil.getJavaDate(Double.parseDouble(text)); // 转为 java.util.Date
                    parse = date.toInstant().atZone(ZoneId.systemDefault()).toLocalTime();
                }
                return parse;
            }
            public static String print(LocalTime object) {  return timeFormatter.format(object);  }
        }

        public static class DateFormatter {
            private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            public static Date parse(String text) {
                Date date;
                try {
                    date = dateFormat.parse(text);
                } catch (Exception e) {
                    // 使用正则提取单引号中的数字
                    Matcher matcher = DATE_PATTERN.matcher(e.getMessage());

                    if (!matcher.find()) {
                        throw new RuntimeException(e);
                    }
                    date = DateUtil.getJavaDate(Double.parseDouble(text)); // 转为 java.util.Date
                }
                return date;
            }
            public static String print(Date date) {  return dateFormat.format(date);  }
        }
        public static class DateTimeFormatterCustom {
            private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            public static Date parse(String text){
                Date date;
                try {
                    date = dateFormat.parse(text);
                } catch (Exception e) {
                    // 使用正则提取单引号中的数字
                    Matcher matcher = DATE_PATTERN.matcher(e.getMessage());

                    if (!matcher.find()) {
                        throw new RuntimeException(e);
                    }
                    date = DateUtil.getJavaDate(Double.parseDouble(text)); // 转为 java.util.Date
                }
                return date;
            }
            public static String print(Date date) {
                return dateFormat.format(date);
            }
        }
    }

    public static void main(String[] args) throws Exception {

        testImport();
//        testExport();

    }

    private static void testImport() throws IOException {

        File file = new File("ExcelUtils-test.xlsx");

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
        map.put("imageUrl", "图片链接");
        map.put("phoneNumber", "手机号");

        List<String> attributeList = new ArrayList<>(map.keySet());



        List<TestObject> testObjectList = importExcel(file, attributeList, TestObject.class);

        for (TestObject testObject : testObjectList) {
            System.out.println(testObject);
        }

    }

    private static void testExport() throws IOException {
        // 输出文件路径
        String outputPath = "ExcelUtils-test" + ".xlsx";

        // 数据
        List<TestObject> testObjectList = new ArrayList<>();
        testObjectList.add(TestObject.builder().phoneNumber("19329651071").imageUrl("http://mms1.baidu.com/it/u=1684950961,555061934&fm=253&app=120&f=JPEG?w=800&h=800").localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(true).build());
        testObjectList.add(TestObject.builder().phoneNumber("19329651072").imageUrl("http://mms0.baidu.com/it/u=1163903759,2895241531&fm=253&app=138&f=JPEG?w=800&h=1066").localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(true).build());
        testObjectList.add(TestObject.builder().phoneNumber("19329651073").imageUrl("http://mms1.baidu.com/it/u=4198565569,2274601556&fm=253&app=138&f=JPEG?w=513&h=500").localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(false).build());
        testObjectList.add(TestObject.builder().phoneNumber("19329651074").imageUrl("http://mms2.baidu.com/it/u=962926323,2652095159&fm=253&app=120&f=JPEG?w=800&h=800").localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(false).build());
        testObjectList.add(TestObject.builder().phoneNumber("19329651075").imageUrl("http://mms2.baidu.com/it/u=3123971159,81579136&fm=253&app=138&f=JPEG?w=500&h=620").localDateTime(LocalDateTime.now()).localDate(LocalDate.now()).localTime(LocalTime.now()).date(new Date()).string("String").integer(666).aFloat(2.5f).aDouble(22.33).aLong(888L).bigDecimal(new BigDecimal("666.888")).aBoolean(true).build());

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
        map.put("imageUrl", "图片链接");
        map.put("phoneNumber", "手机号");


        Workbook workbook = ExcelUtils.exportByImage("导出的表格", 2, 0, "导出的标题",
                testObjectList, map, TestObject.class, false, "imageUrl", 128, 64);

        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
        workbook.write(fileOutputStream);
        workbook.close();

        List<String> attributeList = new ArrayList<>(map.keySet());
//        List<String> attributeList = Arrays.asList("localDateTime", "localDate", "localTime", "date", "string", "integer", "aFloat", "aDouble", "aLong", "bigDecimal", "aBoolean");
        // 测试导入
        List<TestObject> testObjects = importExcel(new File(outputPath), attributeList, TestObject.class);
        System.out.println(testObjects);
    }

}
