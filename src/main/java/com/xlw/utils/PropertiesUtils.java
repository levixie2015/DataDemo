package com.xlw.utils;

import java.io.*;
import java.util.Properties;

/**
 * 读取配置文件工具
 */
public class PropertiesUtils {
    private Properties properties = new Properties();

    public PropertiesUtils(String propertiesFileName) {
        try (InputStream input = new FileInputStream(propertiesFileName)) {
            if (input == null) {
                System.out.println("Sorry, unable to find " + propertiesFileName);
                return;
            }
            // 使用BufferedReader和InputStreamReader来指定UTF-8编码
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(input, "UTF-8"))) {
                // 从BufferedReader加载属性文件
                properties.load(reader);
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }

//        try (InputStream resourceAsStream = new FileInputStream(propertiesFileName)) {
//            properties.load(resourceAsStream);
//            System.out.println("读取配置文件：" + propertiesFileName);
//            for (Map.Entry entry : properties.entrySet()) {
//                System.out.println(entry.getKey() + "--->" + entry.getValue());
//            }
//            System.out.println("读取配置文件结束");
//        } catch (IOException ex) {
//            ex.printStackTrace();
//        }
    }

    public String getProperty(String key) {
        return properties.getProperty(key);
    }

    public static void main(String[] args) {
        //payPlan.sourceExcel=/Users/xieliwei/Desktop/付款计划处理/付款单模板.xlsx
        ///Users/xieliwei/Desktop/付款计划处理/程序/payPlan.txt
        PropertiesUtils config = new PropertiesUtils("/Users/xieliwei/Desktop/付款计划处理/程序/payPlan.txt");
//        PropertiesUtils config = new PropertiesUtils("/payPlan.txt");
        System.out.println("payPlan.sourceExcel: " + config.getProperty("payPlan.sourceExcel"));

    }
}