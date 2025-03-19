package com.hosecloud.demo.vo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * Excel解析结果
 */
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