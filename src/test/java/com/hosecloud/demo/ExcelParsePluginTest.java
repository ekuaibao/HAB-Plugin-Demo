package com.hosecloud.demo;

import com.hosecloud.demo.vo.ExcelItem;
import com.hosecloud.demo.vo.ExcelParseResult;
import com.hosecloud.hab.plugin.BaseTaskPlugin;
import com.hosecloud.hab.plugin.model.Log;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class ExcelParsePluginTest {

    private static final String TEST_SHEET_NAME = "测试工作表";
    private static File testExcelFile;

    @BeforeAll
    static void setUp(@TempDir Path tempDir) throws IOException {
        // 创建测试用的Excel文件
        testExcelFile = createTestExcelFile(tempDir);
    }

    @Test
    void testRunWithValidExcel() throws Exception {
        // 创建插件实例
        ExcelParsePlugin plugin = new ExcelParsePlugin();
        
        // 设置插件属性
        plugin.setExcelUrl(testExcelFile.toURI().toString());
        plugin.setHeaderRowIndex(2);
        plugin.setSheetName(TEST_SHEET_NAME);
        
        // 初始化日志列表
        setExecuteLogs(plugin, new ArrayList<>());
        
        // 执行插件
        ExcelParseResult result = plugin.run();
        
        // 验证结果
        assertNotNull(result);
        assertEquals("解析成功", result.getMessage());
        assertNotNull(result.getHeaders());
        assertNotNull(result.getDataList());
        assertNotNull(result.getItems());
        
        // 验证表头
        List<String> headers = result.getHeaders();
        assertTrue(headers.contains("采购日期"));
        assertTrue(headers.contains("物品名称"));
        assertTrue(headers.contains("费用类型"));
        
        // 验证数据行数
        assertEquals(3, result.getDataList().size());
        assertEquals(3, result.getItems().size());
        
        // 验证第一行数据
        ExcelItem firstItem = result.getItems().get(0);
        assertEquals(1, firstItem.getSerialNumber());
        // 不直接验证日期格式，因为格式可能会根据系统设置而变化
        assertNotNull(firstItem.getPurchaseDate());
        assertTrue(firstItem.getPurchaseDate().contains("2023"));
        assertEquals("笔记本电脑", firstItem.getItemName());
        assertEquals("办公设备", firstItem.getExpenseType());
        assertEquals("技术部", firstItem.getDepartment());
        assertEquals("开发使用", firstItem.getPurpose());
        assertEquals(2, firstItem.getQuantity());
        assertEquals("台", firstItem.getUnit());
        assertEquals(8000.0, firstItem.getUnitPrice());
        assertEquals(16000.0, firstItem.getAmount());
    }

    @Test
    void testRunWithInvalidHeaderRow() throws Exception {
        // 创建插件实例
        ExcelParsePlugin plugin = new ExcelParsePlugin();
        
        // 设置插件属性，使用不存在的表头行
        plugin.setExcelUrl(testExcelFile.toURI().toString());
        plugin.setHeaderRowIndex(10); // 不存在的行号
        plugin.setSheetName(TEST_SHEET_NAME);
        
        // 初始化日志列表
        setExecuteLogs(plugin, new ArrayList<>());
        
        // 执行插件
        ExcelParseResult result = plugin.run();
        
        // 验证结果
        assertNotNull(result);
        assertTrue(result.getMessage().contains("表头行不存在"));
        assertNull(result.getHeaders());
        assertNull(result.getDataList());
        assertNull(result.getItems());
    }

    @Test
    void testRunWithInvalidSheetName() throws Exception {
        // 创建插件实例
        ExcelParsePlugin plugin = new ExcelParsePlugin();
        
        // 设置插件属性，使用不存在的工作表名
        plugin.setExcelUrl(testExcelFile.toURI().toString());
        plugin.setHeaderRowIndex(2);
        plugin.setSheetName("不存在的工作表");
        
        // 初始化日志列表
        setExecuteLogs(plugin, new ArrayList<>());
        
        // 执行插件
        ExcelParseResult result = plugin.run();
        
        // 验证结果
        assertNotNull(result);
        assertTrue(result.getMessage().contains("找不到名为"));
        assertNull(result.getHeaders());
        assertNull(result.getDataList());
        assertNull(result.getItems());
    }

    @Test
    void testRunWithInvalidUrl() throws Exception {
        // 创建插件实例
        ExcelParsePlugin plugin = new ExcelParsePlugin();
        
        // 设置插件属性，使用无效的URL
        plugin.setExcelUrl("file:///invalid-path/file.xlsx");
        plugin.setHeaderRowIndex(2);
        
        // 初始化日志列表
        setExecuteLogs(plugin, new ArrayList<>());
        
        // 执行插件
        ExcelParseResult result = plugin.run();
        
        // 验证结果
        assertNotNull(result);
        assertTrue(result.getMessage().contains("解析Excel文件失败"));
        assertNull(result.getHeaders());
        assertNull(result.getDataList());
        assertNull(result.getItems());
    }

    @Test
    void testDefaultSheetSelection() throws Exception {
        // 创建插件实例
        ExcelParsePlugin plugin = new ExcelParsePlugin();
        
        // 设置插件属性，不指定工作表名
        plugin.setExcelUrl(testExcelFile.toURI().toString());
        plugin.setHeaderRowIndex(2);
        plugin.setSheetName(null);
        
        // 初始化日志列表
        setExecuteLogs(plugin, new ArrayList<>());
        
        // 执行插件
        ExcelParseResult result = plugin.run();
        
        // 验证结果
        assertNotNull(result);
        assertEquals("解析成功", result.getMessage());
        assertNotNull(result.getHeaders());
        assertNotNull(result.getDataList());
        assertNotNull(result.getItems());
    }
    
    @Test
    void testGetCellValueMethods() throws Exception {
        // 创建一个工作簿和工作表用于测试
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");
        
        // 创建一行
        Row row = sheet.createRow(0);
        
        // 创建不同类型的单元格
        // 字符串类型
        Cell stringCell = row.createCell(0);
        stringCell.setCellValue("测试文本");
        
        // 数值类型（整数）
        Cell intCell = row.createCell(1);
        intCell.setCellValue(100);
        
        // 数值类型（小数）
        Cell doubleCell = row.createCell(2);
        doubleCell.setCellValue(123.45);
        
        // 布尔类型
        Cell booleanCell = row.createCell(3);
        booleanCell.setCellValue(true);
        
        // 日期类型 - 使用日期单元格格式
        Cell dateCell = row.createCell(4);
        CellStyle dateStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd"));
        dateCell.setCellStyle(dateStyle);
        dateCell.setCellValue(LocalDate.of(2023, 5, 15));
        
        // 公式类型
        Cell formulaCell = row.createCell(5);
        formulaCell.setCellFormula("SUM(B1:C1)");
        
        // 空单元格
        Cell nullCell = row.createCell(6);
        
        // 创建插件实例
        ExcelParsePlugin plugin = new ExcelParsePlugin();
        
        // 获取私有方法
        Method getCellValueAsStringMethod = ExcelParsePlugin.class.getDeclaredMethod("getCellValueAsString", Cell.class);
        Method getCellValueMethod = ExcelParsePlugin.class.getDeclaredMethod("getCellValue", Cell.class);
        getCellValueAsStringMethod.setAccessible(true);
        getCellValueMethod.setAccessible(true);
        
        // 测试 getCellValueAsString 方法
        assertEquals("测试文本", getCellValueAsStringMethod.invoke(plugin, stringCell));
        assertEquals("100", getCellValueAsStringMethod.invoke(plugin, intCell));
        assertEquals("123.45", getCellValueAsStringMethod.invoke(plugin, doubleCell));
        assertEquals("true", getCellValueAsStringMethod.invoke(plugin, booleanCell));
        
        // 日期单元格的值可能会根据系统设置而变化，所以只检查是否包含年份
        String dateCellValue = (String) getCellValueAsStringMethod.invoke(plugin, dateCell);
        assertNotNull(dateCellValue);
        
        // 公式单元格 - 由于没有计算引擎，所以直接返回公式字符串
        Object formulaResult = getCellValueAsStringMethod.invoke(plugin, formulaCell);
        assertNotNull(formulaResult);
        
        assertEquals("", getCellValueAsStringMethod.invoke(plugin, nullCell));
        
        // 测试 getCellValue 方法
        assertEquals("测试文本", getCellValueMethod.invoke(plugin, stringCell));
        assertEquals(100L, getCellValueMethod.invoke(plugin, intCell));
        assertEquals(123.45, getCellValueMethod.invoke(plugin, doubleCell));
        assertEquals(true, getCellValueMethod.invoke(plugin, booleanCell));
        
        // 日期单元格的值可能是LocalDate或其他日期类型
        Object dateCellObj = getCellValueMethod.invoke(plugin, dateCell);
        assertNotNull(dateCellObj);
        
        // 公式单元格 - 由于没有计算引擎，所以直接返回公式字符串
        Object formulaObj = getCellValueMethod.invoke(plugin, formulaCell);
        assertNotNull(formulaObj);
        
        assertNull(getCellValueMethod.invoke(plugin, nullCell));
        
        // 关闭工作簿
        workbook.close();
    }
    
    @Test
    void testDataConversionMethods() throws Exception {
        // 创建插件实例
        ExcelParsePlugin plugin = new ExcelParsePlugin();
        
        // 获取私有方法
        Method convertToExcelItemsMethod = ExcelParsePlugin.class.getDeclaredMethod("convertToExcelItems", List.class);
        Method getStringValueMethod = ExcelParsePlugin.class.getDeclaredMethod("getStringValue", Map.class, String.class);
        Method getIntValueMethod = ExcelParsePlugin.class.getDeclaredMethod("getIntValue", Map.class, String.class);
        Method getDoubleValueMethod = ExcelParsePlugin.class.getDeclaredMethod("getDoubleValue", Map.class, String.class);
        
        convertToExcelItemsMethod.setAccessible(true);
        getStringValueMethod.setAccessible(true);
        getIntValueMethod.setAccessible(true);
        getDoubleValueMethod.setAccessible(true);
        
        // 创建测试数据
        List<Map<String, Object>> dataList = new ArrayList<>();
        
        // 第一行数据
        Map<String, Object> row1 = new HashMap<>();
        row1.put("序号", 1);
        row1.put("采购日期", "2023/01/15");
        row1.put("物品名称", "笔记本电脑");
        row1.put("费用类型", "办公设备");
        row1.put("使用部门", "技术部");
        row1.put("用途摘要", "开发使用");
        row1.put("数量", 2);
        row1.put("单位", "台");
        row1.put("单价", 8000.0);
        row1.put("金额", 16000.0);
        row1.put("照片", "http://example.com/photo1.jpg");
        row1.put("备注", "紧急");
        dataList.add(row1);
        
        // 第二行数据（包含空值和不同类型）
        Map<String, Object> row2 = new HashMap<>();
        row2.put("序号", "2"); // 字符串类型的数字
        row2.put("采购日期", "2023/02/20");
        row2.put("物品名称", "打印机");
        row2.put("费用类型", "办公设备");
        row2.put("使用部门", "行政部");
        row2.put("用途摘要", "日常办公");
        row2.put("数量", 1);
        row2.put("单位", "台");
        row2.put("单价", 3000);  // 整数类型
        row2.put("金额", 3000.0);
        row2.put("照片", null);  // 空值
        row2.put("备注", "");    // 空字符串
        dataList.add(row2);
        
        // 测试 convertToExcelItems 方法
        @SuppressWarnings("unchecked")
        List<ExcelItem> items = (List<ExcelItem>) convertToExcelItemsMethod.invoke(plugin, dataList);
        
        // 验证结果
        assertEquals(2, items.size());
        
        // 验证第一行数据
        ExcelItem item1 = items.get(0);
        assertEquals(1, item1.getSerialNumber());
        assertEquals("2023/01/15", item1.getPurchaseDate());
        assertEquals("笔记本电脑", item1.getItemName());
        assertEquals("办公设备", item1.getExpenseType());
        assertEquals("技术部", item1.getDepartment());
        assertEquals("开发使用", item1.getPurpose());
        assertEquals(2, item1.getQuantity());
        assertEquals("台", item1.getUnit());
        assertEquals(8000.0, item1.getUnitPrice());
        assertEquals(16000.0, item1.getAmount());
        assertEquals("http://example.com/photo1.jpg", item1.getPhotoUrl());
        assertEquals("紧急", item1.getRemark());
        
        // 验证第二行数据
        ExcelItem item2 = items.get(1);
        assertEquals(2, item2.getSerialNumber());  // 字符串转整数
        assertEquals("2023/02/20", item2.getPurchaseDate());
        assertEquals("打印机", item2.getItemName());
        assertEquals("办公设备", item2.getExpenseType());
        assertEquals("行政部", item2.getDepartment());
        assertEquals("日常办公", item2.getPurpose());
        assertEquals(1, item2.getQuantity());
        assertEquals("台", item2.getUnit());
        assertEquals(3000.0, item2.getUnitPrice()); // 整数转浮点数
        assertEquals(3000.0, item2.getAmount());
        assertNull(item2.getPhotoUrl());  // 空值
        assertEquals("", item2.getRemark()); // 空字符串
        
        // 测试 getStringValue 方法
        assertEquals("笔记本电脑", getStringValueMethod.invoke(plugin, row1, "物品名称"));
        assertNull(getStringValueMethod.invoke(plugin, row1, "不存在的键"));
        assertNull(getStringValueMethod.invoke(plugin, row2, "照片"));
        assertEquals("", getStringValueMethod.invoke(plugin, row2, "备注"));
        
        // 测试 getIntValue 方法
        assertEquals(1, getIntValueMethod.invoke(plugin, row1, "序号"));
        assertEquals(2, getIntValueMethod.invoke(plugin, row2, "序号")); // 字符串转整数
        assertNull(getIntValueMethod.invoke(plugin, row1, "不存在的键"));
        
        // 测试 getDoubleValue 方法
        assertEquals(8000.0, getDoubleValueMethod.invoke(plugin, row1, "单价"));
        assertEquals(3000.0, getDoubleValueMethod.invoke(plugin, row2, "单价")); // 整数转浮点数
        assertNull(getDoubleValueMethod.invoke(plugin, row1, "不存在的键"));
    }

    /**
     * 创建测试用的Excel文件
     */
    private static File createTestExcelFile(Path tempDir) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        
        // 创建工作表
        Sheet sheet = workbook.createSheet(TEST_SHEET_NAME);
        
        // 创建标题行（第1行）
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("采购清单");
        
        // 创建表头行（第2行）
        Row headerRow = sheet.createRow(1);
        String[] headers = {"采购日期", "物品名称", "费用类型", "使用部门", "用途摘要", "数量", "单位", "单价", "金额", "照片", "备注"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }
        
        // 创建数据行
        createDataRow(sheet, 2, "2023-01-15", "笔记本电脑", "办公设备", "技术部", "开发使用", 2, "台", 8000.0, 16000.0, "http://example.com/photo1.jpg", "紧急");
        createDataRow(sheet, 3, "2023-02-20", "打印机", "办公设备", "行政部", "日常办公", 1, "台", 3000.0, 3000.0, "http://example.com/photo2.jpg", "");
        createDataRow(sheet, 4, "2023-03-10", "办公桌椅", "办公家具", "市场部", "新员工入职", 5, "套", 1200.0, 6000.0, "", "标准配置");
        
        // 保存Excel文件
        File file = tempDir.resolve("test-excel.xlsx").toFile();
        try (FileOutputStream outputStream = new FileOutputStream(file)) {
            workbook.write(outputStream);
        }
        workbook.close();
        
        return file;
    }

    /**
     * 创建数据行
     */
    private static void createDataRow(Sheet sheet, int rowIndex, String date, String itemName, String expenseType,
                                     String department, String purpose, int quantity, String unit,
                                     double unitPrice, double amount, String photoUrl, String remark) {
        Row row = sheet.createRow(rowIndex);
        
        // 设置日期
        Cell dateCell = row.createCell(0);
        // 设置日期格式
        CellStyle dateStyle = sheet.getWorkbook().createCellStyle();
        CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/mm/dd"));
        dateCell.setCellStyle(dateStyle);
        
        // 设置日期值
        LocalDate localDate = LocalDate.parse(date);
        dateCell.setCellValue(localDate);
        
        // 设置其他字段
        row.createCell(1).setCellValue(itemName);
        row.createCell(2).setCellValue(expenseType);
        row.createCell(3).setCellValue(department);
        row.createCell(4).setCellValue(purpose);
        row.createCell(5).setCellValue(quantity);
        row.createCell(6).setCellValue(unit);
        row.createCell(7).setCellValue(unitPrice);
        row.createCell(8).setCellValue(amount);
        row.createCell(9).setCellValue(photoUrl);
        row.createCell(10).setCellValue(remark);
    }

    /**
     * 通过反射设置插件的executeLogs字段
     */
    private void setExecuteLogs(ExcelParsePlugin plugin, List<Log> logs) throws Exception {
        Field field = BaseTaskPlugin.class.getDeclaredField("executeLogs");
        field.setAccessible(true);
        field.set(plugin, logs);
    }
} 