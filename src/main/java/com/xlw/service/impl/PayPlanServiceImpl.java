package com.xlw.service.impl;

import cn.hutool.core.util.ObjectUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.listener.PageReadListener;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.xlw.model.PayPlanModel;
import com.xlw.model.PayPlanSupplierModel;
import com.xlw.model.PayPlanTemplateData;
import com.xlw.service.PayPlanService;
import com.xlw.utils.*;
import javafx.util.Pair;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@Data
@AllArgsConstructor
public class PayPlanServiceImpl implements PayPlanService {
    private PropertiesUtils config;

    @Override
    public void parseExcel() {
        String sourceExcelPath = config.getProperty("payPlan.sourceExcel");//待处理的excel源文件
        String additionalExcelPath = config.getProperty("payPlan.additionalExcel");//供应商信息补充文件.若配置则供应商信息优先取此处

        //匹配供应商方式:(0-供应商编码、1-供应商名称).默认供应商编码匹配
        String match = config.getProperty("payPlan.match");
        if (StringUtils.isEmpty(match) || Objects.equals("0", match)) {
            System.out.println("根据供应商编码匹配信息");
        } else {
            System.out.println("根据供应商名称匹配信息");
        }

        List<PayPlanSupplierModel> additionalList = new ArrayList<PayPlanSupplierModel>();
        if (!StringUtils.isEmpty(additionalExcelPath)) {
            System.out.println("已配置【供应商信息补充文件】: " + additionalExcelPath);
            System.out.println("优先从【供应商信息补充文件】取供应商信息");
            //【供应商付款信息表】
            EasyExcel.read(additionalExcelPath, PayPlanSupplierModel.class, new PageReadListener<PayPlanSupplierModel>(dataList -> {
                        for (PayPlanSupplierModel data : dataList) {
                            additionalList.add(data);
                        }
                    })).excelType(ExcelTypeEnum.XLSX) // 指定文件类型为XLSX
                    .sheet(0)// 选择第一个sheet
                    .headRowNumber(1) // 设置从第2行开始读取
                    .doRead();
        } else {
            System.out.println("未配置【供应商信息补充文件");
        }

        String parentPath = Paths.get(sourceExcelPath).getParent().toString();

        //1.获取第一个sheet【付款计划打印版】中的标题
        String contentTitle = getContentTitle(sourceExcelPath);

        //2.获取sheet1、sheet2内容
        Pair<List<PayPlanModel>, List<PayPlanSupplierModel>> sheetContent = getSheetContent(sourceExcelPath);

        //3.获取填充模版数据
        Pair<String, List<PayPlanTemplateData>> pair = getTemplateDataList(contentTitle, sheetContent, additionalList);

        //3.写入excel
        String title = pair.getKey();//sheet标题
        String targetExcelPath = parentPath + File.separator + contentTitle + ".xlsx";
        List<PayPlanTemplateData> templateDataList = pair.getValue();//模版填充数据集
//        EasyExcel.write(targetExcelPath, PayPlanTemplateData.class)
//                .sheet("付款单" + title + "月")
//                .sheetNo(0)
//                .doWrite(templateDataList);

        PayPlanExcelUtils.write(templateDataList, "付款单" + title + "月", targetExcelPath);
        System.out.println("文件处理完成." + targetExcelPath);
    }

    /**
     * 获取第一个sheet【付款计划打印版】中的标题
     *
     * @param sourceExcelPath
     * @return
     */
    private String getContentTitle(String sourceExcelPath) {
        try (FileInputStream fileInputStream = new FileInputStream(sourceExcelPath)) {
            Workbook workbook = WorkbookFactory.create(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);

            Row row = sheet.getRow(0);
            return row.getCell(0).toString();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }

    /**
     * 获取填充模版数据
     *
     * @param contentTitle
     * @param sheetContent
     * @param additionalList 供应商额外补充信息
     * @return
     */
    private Pair<String, List<PayPlanTemplateData>> getTemplateDataList(String contentTitle, Pair<List<PayPlanModel>, List<PayPlanSupplierModel>> sheetContent, List<PayPlanSupplierModel> additionalList) {
        List<PayPlanModel> sheet1 = sheetContent.getKey();//付款计划打印版
        List<PayPlanSupplierModel> sheet2 = sheetContent.getValue();//供应商付款信息表

        PayPlanModel sheet1Title = sheet1.get(0);//获取sheet1标题
        sheet1.remove(0);//删除sheet1标题

        //填充模版数据集合
        List<PayPlanTemplateData> resultList = new ArrayList<PayPlanTemplateData>(sheet1.size());
        for (PayPlanModel item : sheet1) {
            PayPlanTemplateData data = new PayPlanTemplateData();
            data.setTitile(StringUtils.replaceAllBlank(contentTitle));//标题
            data.setSerialNo(StringUtils.replaceAllBlank(item.getSerialNo()));//编号
            data.setSupplierName(StringUtils.replaceAllBlank(item.getSupplierName()));//供应商全称
            data.setSupplierCode(StringUtils.replaceAllBlank(item.getSupplierCode()));//供应商编码
            data.setPayAmount(NumberChineseFormatterUtils.format(StringUtils.replaceAllBlank(item.getPlanMonth()), true, true));//付款总额(中文)
            data.setPaymentMethod(StringUtils.replaceAllBlank(item.getPaymentMethod()));//付款方式
            data.setAmount(StringUtils.replaceAllBlank(item.getPlanMonth()));
            data.setPayer("采购中心");//付款单位
            //data.setOperator();//经办人
            //data.setFinance();//财务
            setSupplierInfo(sheet2, data, additionalList);
            resultList.add(data);
        }
        return new Pair<>(getTitle(sheet1Title), resultList);
    }

    /**
     * 设置供应商信息
     *
     * @param sheet2         供应商信息
     * @param data           最终导出数据
     * @param additionalList 供应商额外补充信息
     */
    private void setSupplierInfo(List<PayPlanSupplierModel> sheet2, PayPlanTemplateData data, List<PayPlanSupplierModel> additionalList) {
        for (PayPlanSupplierModel item : sheet2) {
            //匹配供应商方式:(0-供应商编码、1-供应商名称).默认供应商编码匹配
            String match = config.getProperty("payPlan.match");
            if (StringUtils.isEmpty(match)) {
                match = "0";
            }
            if (Objects.equals("0", match)) {
                if (Objects.equals(data.getSupplierCode(), StringUtils.replaceAllBlank(item.getSupplierCode()))) {
                    doSetData(data, item, additionalList, match);
                }
            } else {
                if (Objects.equals(data.getSupplierName(), StringUtils.replaceAllBlank(item.getSupplierName()))) {
                    doSetData(data, item, additionalList, match);
                }
            }

        }
    }

    /**
     * 设置数据
     *
     * @param data  最终导出数据
     * @param item  供应商信息
     * @param match 匹配供应商方式:(0-供应商编码、1-供应商名称).默认供应商编码匹配
     */
    private void doSetData(PayPlanTemplateData data, PayPlanSupplierModel item, List<PayPlanSupplierModel> additionalList, String match) {
        PayPlanSupplierModel additionalItem = null;
        if (!ObjectUtil.isEmpty(additionalList)) {
            if (Objects.equals("0", match)) {
                additionalItem = additionalList.stream().filter(entity -> Objects.equals(StringUtils.replaceAllBlank(entity.getSupplierCode()), data.getSupplierCode())).findFirst().orElse(null);
            } else {
                additionalItem = additionalList.stream().filter(entity -> Objects.equals(StringUtils.replaceAllBlank(entity.getSupplierName()), data.getSupplierName())).findFirst().orElse(null);
            }
        }

        if (additionalItem != null) {
            item = additionalItem;
        }

        data.setOpeningBank(StringUtils.replaceAllBlank(item.getOpeningBank()));//开户行
        data.setAccountNo(StringUtils.replaceAllBlank(item.getAccountNo()));//账号
        data.setOpeningBankNo(StringUtils.replaceAllBlank(item.getOpeningBankNo()));//行号
    }

    /**
     * 获取sheet标题
     *
     * @param sheet1Title
     * @return 月份。如：9
     */
    private String getTitle(PayPlanModel sheet1Title) {
        List<String> integers = NumberExtractor.extractIntegers(sheet1Title.getPlanMonth());
        return integers.get(0);
    }

    /**
     * 获取sheet1、sheet2内容
     *
     * @param sourceExcelPath 待处理的excel源文件
     * @return Pair<sheet1, sheet2>
     */
    private Pair<List<PayPlanModel>, List<PayPlanSupplierModel>> getSheetContent(String sourceExcelPath) {
        //sheet1: 【付款计划打印版】
        List<PayPlanModel> sheet1List = new ArrayList<PayPlanModel>();
        EasyExcel.read(sourceExcelPath, PayPlanModel.class, new PageReadListener<PayPlanModel>(dataList -> {
                    for (PayPlanModel data : dataList) {
                        sheet1List.add(data);
                    }
                })).excelType(ExcelTypeEnum.XLSX) // 指定文件类型为XLSX
                .sheet(0)// 选择第一个sheet
                //.headRowNumber(3) // 设置从第四行开始读取
                .doRead();

        //sheet2: 【供应商付款信息表】
        List<PayPlanSupplierModel> sheet2List = new ArrayList<PayPlanSupplierModel>();
        EasyExcel.read(sourceExcelPath, PayPlanSupplierModel.class, new PageReadListener<PayPlanSupplierModel>(dataList -> {
                    for (PayPlanSupplierModel data : dataList) {
                        sheet2List.add(data);
                    }
                })).excelType(ExcelTypeEnum.XLSX) // 指定文件类型为XLSX
                .sheet(1)// 选择第一个sheet
                .headRowNumber(1) // 设置从第2行开始读取
                .doRead();

        return new Pair<>(sheet1List, sheet2List);
    }
}
