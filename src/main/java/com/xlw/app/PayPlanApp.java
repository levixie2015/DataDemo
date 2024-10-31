package com.xlw.app;

import com.xlw.service.PayPlanService;
import com.xlw.service.impl.PayPlanServiceImpl;
import com.xlw.utils.PropertiesUtils;

import java.io.File;

/**
 * 付款计划入口
 */
public class PayPlanApp {
    public static void main(String[] args) {

        String currentPath = System.getProperty("user.dir");
        System.out.println("当前路径：" + currentPath);
        String configFilePath = currentPath + File.separator + "payPlan.txt";
        System.out.println("配置文件路径：" + configFilePath);

        //读取配置文件内容
        PropertiesUtils config = new PropertiesUtils(configFilePath);

        //解析excel,把结果保存至excel
        PayPlanService payPlanService = new PayPlanServiceImpl(config);
        payPlanService.parseExcel();
    }
}
