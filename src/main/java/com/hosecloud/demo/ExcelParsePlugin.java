package com.hosecloud.demo;

import com.hosecloud.demo.vo.ExcelItem;
import com.hosecloud.demo.vo.ExcelParseResult;
import com.hosecloud.hab.plugin.BaseTaskPlugin;
import com.hosecloud.hab.plugin.annotation.Execute;
import com.hosecloud.hab.plugin.annotation.JsonSchemaDefinition;
import com.hosecloud.hab.plugin.annotation.JsonSchemaProperty;
import com.hosecloud.hab.plugin.model.Log;
import lombok.Setter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.pf4j.Extension;

import java.io.InputStream;
import java.net.URL;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Setter
@Extension
@JsonSchemaDefinition(
        title = "Excel解析节点",
        description = "解析Excel文件并以列表形式输出数据，支持指定表头行"
)
public class ExcelParsePlugin extends BaseTaskPlugin {

    @JsonSchemaProperty(
            title = "Excel文件链接",
            description = "需要解析的Excel文件URL链接",
            required = true,
            example = "https://example.com/sample.xlsx"
    )
    private String excelUrl;

    @JsonSchemaProperty(
            title = "表头行号",
            description = "指定表头所在的行号（从1开始计数），该行之前的数据将被忽略",
            required = true,
            example = "2"
    )
    private Integer headerRowIndex;

    @JsonSchemaProperty(
            title = "工作表名称",
            description = "指定要解析的工作表名称，不指定则默认解析第一个工作表",
            required = false,
            example = "Sheet1"
    )
    private String sheetName;

    @Override
    public String getName() {
        return "task-dynamic-excel-parse";
    }

    @Execute(
            description = "执行Excel解析操作",
            outputClass = ExcelParseResult.class
    )
    public ExcelParseResult run() {
        try {
            // 下载Excel文件
            URL url = new URL(excelUrl);
            try (InputStream inputStream = url.openStream()) {
                // 创建工作簿
                Workbook workbook = new XSSFWorkbook(inputStream);
                
                // 获取工作表
                Sheet sheet;
                if (sheetName != null && !sheetName.trim().isEmpty()) {
                    sheet = workbook.getSheet(sheetName);
                    if (sheet == null) {
                        executeLogs.add(Log.failure("找不到名为 '" + sheetName + "' 的工作表"));
                        return new ExcelParseResult("找不到名为 '" + sheetName + "' 的工作表");
                    }
                } else {
                    sheet = workbook.getSheetAt(0);
                }
                
                // 获取表头行
                Row headerRow = sheet.getRow(headerRowIndex - 1);
                if (headerRow == null) {
                    executeLogs.add(Log.failure("表头行不存在，请检查表头行号是否正确"));
                    return new ExcelParseResult("表头行不存在，请检查表头行号是否正确");
                }
                
                // 解析表头
                List<String> headers = new ArrayList<>();
                int lastCellNum = headerRow.getLastCellNum();
                for (int i = 0; i < lastCellNum; i++) {
                    Cell cell = headerRow.getCell(i);
                    String headerName = getCellValueAsString(cell);
                    // 如果表头为空，使用列索引作为表头
                    if (headerName == null || headerName.trim().isEmpty()) {
                        headerName = "Column" + (i + 1);
                    }
                    headers.add(headerName);
                }
                
                // 解析数据行
                List<Map<String, Object>> dataList = new ArrayList<>();
                for (int i = headerRowIndex; i <= sheet.getLastRowNum(); i++) {
                    Row dataRow = sheet.getRow(i);
                    if (dataRow == null) continue;
                    
                    Map<String, Object> rowData = new LinkedHashMap<>();
                    boolean hasData = false;
                    
                    // 添加序号字段
                    rowData.put("序号", i - headerRowIndex + 1);
                    
                    for (int j = 0; j < headers.size(); j++) {
                        Cell cell = dataRow.getCell(j);
                        String headerName = headers.get(j);
                        
                        if (cell != null) {
                            Object cellValue = getCellValue(cell);
                            
                            // 特殊处理日期格式
                            if (headerName.contains("日期") && cellValue instanceof LocalDate) {
                                LocalDate date = (LocalDate) cellValue;
                                cellValue = date.format(DateTimeFormatter.ofPattern("yyyy/MM/dd"));
                            }
                            
                            // 特殊处理数字格式
                            if (cellValue instanceof Double) {
                                Double numValue = (Double) cellValue;
                                // 如果是整数，转换为整数显示
                                if (numValue == Math.floor(numValue)) {
                                    cellValue = numValue.longValue();
                                }
                            }
                            
                            rowData.put(headerName, cellValue);
                            if (cellValue != null && !cellValue.toString().trim().isEmpty()) {
                                hasData = true;
                            }
                        } else {
                            rowData.put(headerName, null);
                        }
                    }
                    
                    // 只添加非空行
                    if (hasData) {
                        dataList.add(rowData);
                    }
                }
                
                // 关闭工作簿
                workbook.close();
                
                // 将数据转换为ExcelItem对象列表
                List<ExcelItem> items = convertToExcelItems(dataList);
                
                // 记录日志
                executeLogs.add(Log.success("成功解析Excel文件，共解析" + dataList.size() + "行数据"));
                
                // 返回结果
                ExcelParseResult result = new ExcelParseResult();
                result.setMessage("解析成功");
                result.setHeaders(headers);
                result.setDataList(dataList);
                result.setItems(items);
                return result;
            }
        } catch (Exception e) {
            executeLogs.add(Log.failure("解析Excel文件失败: " + e.getMessage()));
            return new ExcelParseResult("解析Excel文件失败: " + e.getMessage());
        }
    }
    
    /**
     * 将Map列表转换为ExcelItem对象列表
     */
    private List<ExcelItem> convertToExcelItems(List<Map<String, Object>> dataList) {
        return dataList.stream().map(row -> {
            ExcelItem item = new ExcelItem();
            
            // 设置序号
            item.setSerialNumber(getIntValue(row, "序号"));
            
            // 设置采购日期
            item.setPurchaseDate(getStringValue(row, "采购日期"));
            
            // 设置物品名称
            item.setItemName(getStringValue(row, "物品名称"));
            
            // 设置费用类型
            item.setExpenseType(getStringValue(row, "费用类型"));
            
            // 设置使用部门
            item.setDepartment(getStringValue(row, "使用部门"));
            
            // 设置用途摘要
            item.setPurpose(getStringValue(row, "用途摘要"));
            
            // 设置数量
            item.setQuantity(getIntValue(row, "数量"));
            
            // 设置单位
            item.setUnit(getStringValue(row, "单位"));
            
            // 设置单价
            item.setUnitPrice(getDoubleValue(row, "单价"));
            
            // 设置金额
            item.setAmount(getDoubleValue(row, "金额"));
            
            // 设置照片URL
            item.setPhotoUrl(getStringValue(row, "照片"));
            
            // 设置备注
            item.setRemark(getStringValue(row, "备注"));
            
            return item;
        }).collect(Collectors.toList());
    }
    
    /**
     * 从Map中获取字符串值
     */
    private String getStringValue(Map<String, Object> row, String key) {
        Object value = row.get(key);
        return value != null ? value.toString() : null;
    }
    
    /**
     * 从Map中获取整数值
     */
    private Integer getIntValue(Map<String, Object> row, String key) {
        Object value = row.get(key);
        if (value == null) return null;
        
        if (value instanceof Number) {
            return ((Number) value).intValue();
        }
        
        try {
            return Integer.parseInt(value.toString());
        } catch (NumberFormatException e) {
            return null;
        }
    }
    
    /**
     * 从Map中获取浮点数值
     */
    private Double getDoubleValue(Map<String, Object> row, String key) {
        Object value = row.get(key);
        if (value == null) return null;
        
        if (value instanceof Number) {
            return ((Number) value).doubleValue();
        }
        
        try {
            return Double.parseDouble(value.toString());
        } catch (NumberFormatException e) {
            return null;
        }
    }
    
    /**
     * 获取单元格的值（字符串形式）
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toString();
                } else {
                    double value = cell.getNumericCellValue();
                    // 如果是整数，返回整数形式
                    if (value == Math.floor(value)) {
                        return String.valueOf((long) value);
                    }
                    // 避免数值显示为科学计数法
                    return String.valueOf(value);
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return String.valueOf((long) value);
                    }
                    return String.valueOf(value);
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ex) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }
    
    /**
     * 获取单元格的值（保留原始类型）
     */
    private Object getCellValue(Cell cell) {
        if (cell == null) return null;
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toLocalDate();
                } else {
                    double value = cell.getNumericCellValue();
                    // 如果是整数，返回整数类型
                    if (value == Math.floor(value)) {
                        return (long) value;
                    }
                    return value;
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                try {
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return (long) value;
                    }
                    return value;
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ex) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return null;
        }
    }
} 