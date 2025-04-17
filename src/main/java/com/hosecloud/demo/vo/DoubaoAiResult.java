package com.hosecloud.demo.vo;

import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.AllArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class DoubaoAiResult {
    /**
     * 响应状态信息
     */
    private String message;
    
    /**
     * AI返回的文本内容
     */
    private String content;
    
    /**
     * 构造函数
     * @param message 状态信息
     */
    public DoubaoAiResult(String message) {
        this.message = message;
    }
} 