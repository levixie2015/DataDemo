package com.xlw.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.io.Serializable;

/**
 * 供应商付款信息表
 */
@Data
public class PayPlanSupplierModel implements Serializable {
    /**
     * 供应商编码
     */
    @ExcelProperty(value = "供应商编码", index = 0)
    private String supplierCode;
    /**
     * 供应商全称
     */
    @ExcelProperty(value = "供应商全称", index = 1)
    private String supplierName;
    /**
     * 开户行
     */
    @ExcelProperty(value = "开户行", index = 2)
    private String openingBank;
    /**
     * 账号
     */
    @ExcelProperty(value = "账号", index = 3)
    private String accountNo;
    /**
     * 开户行号
     */
    @ExcelProperty(value = "开户行号", index = 4)
    private String openingBankNo;
}