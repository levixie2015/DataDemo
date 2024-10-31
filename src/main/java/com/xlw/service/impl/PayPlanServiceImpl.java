package com.xlw.service.impl;

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
        String parentPath = Paths.get(sourceExcelPath).getParent().toString();

        //1.获取第一个sheet【付款计划打印版】中的标题
        String contentTitle = getContentTitle(sourceExcelPath);

        //2.获取sheet1、sheet2内容
        Pair<List<PayPlanModel>, List<PayPlanSupplierModel>> sheetContent = getSheetContent(sourceExcelPath);

        //3.获取填充模版数据
        Pair<String, List<PayPlanTemplateData>> pair = getTemplateDataList(contentTitle, sheetContent);

        //3.写入excel
        String title = pair.getKey();//sheet标题
        String targetExcelPath = parentPath + File.separator + contentTitle + ".xlsx";
        List<PayPlanTemplateData> templateDataList = pair.getValue();//模版填充数据集
//        EasyExcel.write(targetExcelPath, PayPlanTemplateData.class)
//                .sheet("付款单" + title + "月")
//                .sheetNo(0)
//                .doWrite(templateDataList);

        PayPlanExcelUtils.write(templateDataList, "付款单" + title + "月", targetExcelPath);
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
     */
    private Pair<String, List<PayPlanTemplateData>> getTemplateDataList(String contentTitle, Pair<List<PayPlanModel>, List<PayPlanSupplierModel>> sheetContent) {
        List<PayPlanModel> sheet1 = sheetContent.getKey();//付款计划打印版
        List<PayPlanSupplierModel> sheet2 = sheetContent.getValue();//供应商付款信息表

        PayPlanModel sheet1Title = sheet1.get(0);//获取sheet1标题
        sheet1.remove(0);//删除sheet1标题

        //填充模版数据集合
        List<PayPlanTemplateData> resultList = new ArrayList<PayPlanTemplateData>(sheet1.size());
        for (PayPlanModel item : sheet1) {
            PayPlanTemplateData data = new PayPlanTemplateData();
            data.setTitile(contentTitle);//标题
            data.setSerialNo(item.getSerialNo());//编号
            data.setSupplierName(item.getSupplierName());//供应商全称
            data.setSupplierCode(item.getSupplierCode());//供应商编码
            data.setPayAmount(NumberChineseFormatterUtils.format(item.getPlanMonth(), true, true));//付款总额(中文)
            data.setPaymentMethod(item.getPaymentMethod());//付款方式
            data.setAmount(StringUtils.replaceAllBlank(item.getPlanMonth()));
            //data.setPayer("");//付款单位
            //data.setOperator();//经办人
            //data.setFinance();//财务
            setSupplierInfo(sheet2, data);
            resultList.add(data);
        }
        return new Pair<>(getTitle(sheet1Title), resultList);
    }

    /**
     * 设置供应商信息
     */
    private void setSupplierInfo(List<PayPlanSupplierModel> sheet2, PayPlanTemplateData data) {
        for (PayPlanSupplierModel item : sheet2) {
            //根据供应商编码查找信息
            if (Objects.equals(data.getSupplierCode(), item.getSupplierCode())) {
                data.setOpeningBank(item.getOpeningBank());//开户行
                data.setAccountNo(item.getAccountNo());//账号
                data.setOpeningBankNo(item.getOpeningBankNo());//行号
            }
        }
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
