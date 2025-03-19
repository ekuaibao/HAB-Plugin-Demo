# 插件开发指南

## 简介

本文档介绍如何通过`ExcelParsePlugin`示例来开发一个自定义插件。该插件演示了如何解析Excel文件并处理其中的数据。

## 插件开发流程

### 1. 创建插件类

插件必须继承`BaseTaskPlugin`基类并使用`@Extension`注解标记。如示例中的：

```java
@Setter
@Extension
@JsonSchemaDefinition(
        title = "Excel解析节点",
        description = "解析Excel文件并以列表形式输出数据，支持指定表头行"
)
public class ExcelParsePlugin extends BaseTaskPlugin {
    // 插件内容
}
```

### 2. 定义插件属性

使用`@JsonSchemaProperty`注解定义插件的配置属性：

```java
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
```

每个属性可以配置：
- `title`: 属性的显示名称
- `description`: 属性的详细描述
- `required`: 是否必填
- `example`: 示例值

### 3. 定义输出模型

为插件定义规范的输出模型，例如我们的示例中定义了两个类：

#### ExcelParseResult（结果类）

```java
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelParseResult {
    /**
     * 解析状态信息
     */
    private String message;
    
    /**
     * 表头列表
     */
    private List<String> headers;
    
    /**
     * 数据列表，每一行是一个Map，键为表头，值为单元格内容
     */
    private List<Map<String, Object>> dataList;
    
    /**
     * 格式化后的数据项列表
     */
    private List<ExcelItem> items;
    
    /**
     * 构造函数
     * @param message 状态信息
     */
    public ExcelParseResult(String message) {
        this.message = message;
    }
}
```

#### ExcelItem（数据项类）

```java
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelItem {
    /**
     * 序号
     */
    private Integer serialNumber;
    
    /**
     * 采购日期
     */
    private String purchaseDate;
    
    /**
     * 物品名称
     */
    private String itemName;
    
    // 更多字段...
}
```

### 4. 实现插件方法

#### 定义插件名称

覆盖`getName()`方法定义插件唯一标识：

```java
@Override
public String getName() {
    return "task-dynamic-excel-parse";
}
```

#### 实现执行方法

使用`@Execute`注解定义插件的执行方法：

```java
@Execute(
        description = "执行Excel解析操作",
        outputClass = ExcelParseResult.class
)
public ExcelParseResult run() {
    try {
        // 1. 下载Excel文件
        URL url = new URL(excelUrl);
        try (InputStream inputStream = url.openStream()) {
            // 2. 创建工作簿
            Workbook workbook = new XSSFWorkbook(inputStream);
            
            // 3. 获取工作表
            Sheet sheet = getTargetSheet(workbook);
            
            // 4. 解析表头和数据
            List<String> headers = parseHeaders(sheet);
            List<Map<String, Object>> dataList = parseData(sheet, headers);
            
            // 5. 转换为结构化对象
            List<ExcelItem> items = convertToExcelItems(dataList);
            
            // 6. 记录日志并返回结果
            executeLogs.add(Log.success("成功解析Excel文件，共解析" + dataList.size() + "行数据"));
            
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
```

### 5. 日志记录

插件中可以使用`executeLogs`记录执行日志：

```java
executeLogs.add(Log.success("成功解析Excel文件，共解析" + dataList.size() + "行数据"));
// 或
executeLogs.add(Log.failure("解析Excel文件失败: " + e.getMessage()));
```

### 6. 处理异常情况

确保插件能够优雅地处理各种异常情况：

```java
try {
    // 业务逻辑
} catch (Exception e) {
    // 记录错误日志
    executeLogs.add(Log.failure("发生错误: " + e.getMessage()));
    // 返回错误结果
    return new ExcelParseResult("发生错误: " + e.getMessage());
}
```

## 插件项目结构

一个完整的插件项目通常包含以下结构：

```
src/
├── main/
│   ├── java/
│   │   └── com/
│   │       └── company/
│   │           └── plugin/
│   │               ├── MyPlugin.java          // 主插件类
│   │               ├── service/               // 服务层
│   │               │   └── BusinessService.java
│   │               └── vo/                    // 值对象
│   │                   ├── InputModel.java    // 输入模型
│   │                   └── ResultModel.java   // 结果模型
│   └── resources/
│       └── META-INF/
│           └── extensions.idx                 // 插件扩展点索引
└── test/                                     // 单元测试
    └── java/
        └── com/
            └── company/
                └── plugin/
                    └── MyPluginTest.java
```

## 测试插件

1. 编译插件项目
2. 将编译后的JAR包放入平台的插件目录
3. 重启服务或触发插件热加载
4. 在工作流编辑器中使用该插件

## 插件示例说明

`ExcelParsePlugin`示例实现了以下功能：

1. 从URL下载Excel文件
2. 根据配置选择工作表
3. 解析表头和数据行
4. 处理不同类型的单元格值（文本、数字、日期等）
5. 将解析结果转换为结构化对象
6. 记录执行日志并返回结果

## 核心方法说明

### 获取工作表

```java
private Sheet getTargetSheet(Workbook workbook) {
    if (sheetName != null && !sheetName.trim().isEmpty()) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            executeLogs.add(Log.failure("找不到名为 '" + sheetName + "' 的工作表"));
            throw new RuntimeException("找不到名为 '" + sheetName + "' 的工作表");
        }
        return sheet;
    } else {
        return workbook.getSheetAt(0);
    }
}
```

### 处理单元格值

```java
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
        // 其他类型处理...
    }
}
```

## 插件开发最佳实践

1. 为插件提供清晰的描述和属性说明
2. 处理各种异常情况，确保插件稳定性
3. 提供合适的日志记录，方便排查问题
4. 返回结构化数据，便于后续节点处理
5. 遵循单一职责原则，每个插件只做一件事
6. 添加必要的注释说明复杂逻辑
7. 代码模块化，将业务逻辑拆分为多个方法
8. 使用VO类规范化数据结构

## 常见问题

1. **插件无法加载**：检查插件类上是否添加了`@Extension`注解
2. **属性无法配置**：检查`@JsonSchemaProperty`注解配置是否正确
3. **执行失败**：查看日志，确保异常处理正确
4. **返回值不正确**：确保`@Execute`注解中指定了正确的`outputClass`

## 参考资料

- [PF4J插件框架文档](https://pf4j.org/)
- [Apache POI文档](https://poi.apache.org/)
- [Lombok文档](https://projectlombok.org/) 