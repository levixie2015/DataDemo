package com.xlw.app;

import com.xlw.service.PayPlanService;
import com.xlw.service.impl.PayPlanServiceImpl;
import com.xlw.utils.PropertiesUtils;

/**
 * 付款计划入口
 */
public class PayPlanApp {
    public static void main(String[] args) {
        String configFilePath = args.length > 0 ? args[0] : "payPlan.properties";

        //读取配置文件内容
        PropertiesUtils config = new PropertiesUtils(configFilePath);

        //解析excel,把结果保存至excel
        PayPlanService payPlanService = new PayPlanServiceImpl(config);
        payPlanService.parseExcel();
    }
}
