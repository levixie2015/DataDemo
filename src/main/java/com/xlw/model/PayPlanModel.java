package com.xlw.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.io.Serializable;

/**
 * 付款计划打印版
 */
@Data
public class PayPlanModel implements Serializable {
    /**
     * 序号
     */
    @ExcelProperty(value = "序号", index = 0)
    private String serialNo;
    /**
     * 供应商编码
     */
    @ExcelProperty(value = "供应商编码", index = 1)
    private String supplierCode;
    /**
     * 供应商全称
     */
    @ExcelProperty(value = "供应商全称", index = 2)
    private String supplierName;
    /**
     * 采购X月付款计划
     */
    @ExcelProperty(value = "采购X月付款计划", index = 3)
    private String planMonth;
    /**
     * 账期
     */
    @ExcelProperty(value = "账期", index = 4)
    private String accountPeriod;
    /**
     * 付款方式
     */
    @ExcelProperty(value = "付款方式", index = 5)
    private String paymentMethod;
    /**
     * 采购员
     */
    @ExcelProperty(value = "采购员", index = 6)
    private String purchaser;
}
