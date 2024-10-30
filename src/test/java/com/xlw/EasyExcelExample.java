package com.xlw;

import cn.hutool.core.bean.BeanUtil;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

public class EasyExcelExample {
    public static void main(String[] args) {
        // 设置Excel文件路径
        String fileName = "/Users/xieliwei/Desktop/付款计划处理/材料付款计划.xlsx";
//        String templateFileName = "/Users/xieliwei/Desktop/付款计划处理/付款单模板.xlsx";

        // 创建数据模型
        List<PaymentPlanDemo> dataList = initDataList();
//        ExcelUtils.write(dataList, "付款单9月", fileName);
    }

    private static List<PaymentPlanDemo> initDataList() {
        List<PaymentPlanDemo> dataList = new ArrayList<PaymentPlanDemo>();
        PaymentPlanDemo data = new PaymentPlanDemo();
        data.setSupplierName("武汉xxx有限公司");
        data.setOpeningBank("招商银行武汉经济技术开发区支行");
        data.setSupplierCode("06023");
        data.setAccountNo("127910689210401");
        data.setPayAmount("壹拾肆万元整");
        data.setOpeningBankNo("308521015209");
        data.setPaymentMethod("迪链");
        data.setAmount("140,000");
        data.setPayer("采购中心");
        data.setOperator("");
        data.setFinance("");
        dataList.add(data);

        PaymentPlanDemo data2 = new PaymentPlanDemo();
        BeanUtil.copyProperties(data, data2);
        dataList.add(data2);

        PaymentPlanDemo data3 = new PaymentPlanDemo();
        BeanUtil.copyProperties(data, data3);
        dataList.add(data3);

        PaymentPlanDemo data4 = new PaymentPlanDemo();
        BeanUtil.copyProperties(data, data4);
        dataList.add(data4);

        PaymentPlanDemo data5 = new PaymentPlanDemo();
        BeanUtil.copyProperties(data, data5);
        dataList.add(data5);

        PaymentPlanDemo data6 = new PaymentPlanDemo();
        BeanUtil.copyProperties(data, data6);
        dataList.add(data6);

        PaymentPlanDemo data7 = new PaymentPlanDemo();
        BeanUtil.copyProperties(data, data7);
        dataList.add(data7);

        return dataList;
    }
}

@Data
class PaymentPlanDemo {
    /**
     * 标题
     */
    private String titile = "2024年9月材料付款计划";
    /**
     * 编号
     */
    private String serialNo = "1";
    /**
     * 供应商全称
     */
    private String supplierName;
    /**
     * 开户行
     */
    private String openingBank;
    /**
     * 供应商编码
     */
    private String supplierCode;
    /**
     * 账号
     */
    private String accountNo;
    /**
     * 付款总额
     */
    private String payAmount;
    /**
     * 行号
     */
    private String openingBankNo;
    /**
     * 付款方式
     */
    private String paymentMethod;
    /**
     * 付款方式
     */
    private String amount;
    /**
     * 付款单位
     */
    private String payer;
    /**
     * 经办人
     */
    private String operator;
    /**
     * 财务
     */
    private String finance;
}