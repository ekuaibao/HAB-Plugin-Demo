package com.hosecloud.demo.vo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Excel数据项，对应Excel中的一行数据
 */
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
    
    /**
     * 费用类型
     */
    private String expenseType;
    
    /**
     * 使用部门
     */
    private String department;
    
    /**
     * 用途摘要
     */
    private String purpose;
    
    /**
     * 数量
     */
    private Integer quantity;
    
    /**
     * 单位
     */
    private String unit;
    
    /**
     * 单价
     */
    private Double unitPrice;
    
    /**
     * 金额
     */
    private Double amount;
    
    /**
     * 照片URL
     */
    private String photoUrl;
    
    /**
     * 备注
     */
    private String remark;
} 