package com.xlw.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.io.Serializable;

@Data
public class PayPlanTemplateData implements Serializable {
    /**
     * 标题
     */
    private String titile;
    /**
     * 编号
     */
    private String serialNo;
    /**
     * 供应商全称
     */
    @ExcelProperty(value = "供应商全称", index = 0)
    private String supplierName;
    /**
     * 开户行
     */
    @ExcelProperty(value = "开户行", index = 1)
    private String openingBank;
    /**
     * 供应商编码
     */
    @ExcelProperty(value = "供应商编码", index = 2)
    private String supplierCode;
    /**
     * 账号
     */
    @ExcelProperty(value = "账号", index = 3)
    private String accountNo;
    /**
     * 付款总额（大写）
     */
    @ExcelProperty(value = "付款总额（大写）", index = 4)
    private String payAmount;
    /**
     * 行号
     */
    @ExcelProperty(value = "行号", index = 5)
    private String openingBankNo;
    /**
     * 付款方式
     */
    @ExcelProperty(value = "付款方式", index = 6)
    private String paymentMethod;
    /**
     * 金额（数字）
     */
    private String amount;
    /**
     * 付款单位
     */
    @ExcelProperty(value = "付款单位", index = 7)
    private String payer;
    /**
     * 经办人
     */
    @ExcelProperty(value = "经办人", index = 8)
    private String operator;
    /**
     * 财务
     */
    @ExcelProperty(value = "财务", index = 9)
    private String finance;
}