package com.xlw.utils;

import java.text.DecimalFormat;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MoneyUtils {

    /**
     * 金额千分位格式化
     *
     * @param text
     * @return
     */
    public static String fmtMicrometer(String text) {
        DecimalFormat df = null;
        if (text.indexOf(".") > 0) {
            if (text.length() - text.indexOf(".") - 1 == 0) {
                df = new DecimalFormat("###,##0.");
            } else if (text.length() - text.indexOf(".") - 1 == 1) {
                df = new DecimalFormat("###,##0.00");
            } else {
                df = new DecimalFormat("###,##0.00");
            }
        } else {
            df = new DecimalFormat("###,##0");
        }
        double number = 0.0;
        try {
            number = Double.parseDouble(text);
        } catch (Exception e) {
            e.printStackTrace();
            number = 0.0;
        }
        return df.format(number);
    }

    /**
     * 将金额（整数部分等于或少于 12 位，小数部分 2 位）转换为中文大写形式
     *
     * @param amount
     * @return
     * @throws IllegalArgumentException
     */
    public static String convertChineseAmount(String amount) throws IllegalArgumentException {
        amount = fmtMicrometer(amount);
        // 去掉分隔符
        amount = amount.replace(",", "");
        // 验证金额正确性
        if ("0.00".equals(amount)) {
            throw new IllegalArgumentException("金额不能为零.");
        }
        // 正则匹配整数和小数部分
        String regex = "^(?<integer>\\d+)(\\.(?<fraction>\\d{1,2}))?$";
        // 匹配器
        Matcher matcher = Pattern.compile(regex).matcher(amount);
        if (!matcher.find()) {
            throw new IllegalArgumentException("输入金额有误.");
        }
        // 整数部分
        String integer = matcher.group("integer");
        // 小数部分
        String fraction = matcher.group("fraction");
        String result = "";
        if (!"0".equals(integer)) {
            // 整数部分
            result += integerToRmb(integer) + "元";
        }
        if ("00".equals(fraction)) {
            // 添加[整]
            result += "整";
        } else if (fraction != null && !"0".equals(integer)) {
            // 小数部分
            result += fractionToRmb(fraction);
        } else if (fraction != null && "0".equals(integer)) {
            // 整数为0时，小数部分
            result += fractionToRmb(fraction.substring(1));
        }
        return result;
    }

    /**
     * 将金额小数部分转换为中文大写
     *
     * @param fraction
     * @return
     */
    private static String fractionToRmb(String fraction) {
        // 角
        char jiao = fraction.charAt(0);
        // 分
        char fen = fraction.charAt(1);
        return (RMB_NUMS[jiao - '0'] + (jiao > '0' ? "角" : ""))
                + (fen > '0' ? RMB_NUMS[fen - '0'] + "分" : "");
    }

    /**
     * 将金额整数部分转换为中文大写
     *
     * @param integer
     * @return
     */
    private static String integerToRmb(String integer) {
        StringBuilder buffer = new StringBuilder();
        // 从个位数开始转换
        int i, j;
        for (i = integer.length() - 1, j = 0; i >= 0; i--, j++) {
            char n = integer.charAt(i);
            if (n == '0') {
                // 当 n 是 0 且 n 的右边一位不是 0 时，插入[零]
                if (i < integer.length() - 1 && integer.charAt(i + 1) != '0') {
                    buffer.append("零");
                }
                // 插入[万]或者[亿]
                if (j % 4 == 0) {
                    if (i > 0 && integer.charAt(i - 1) != '0' || i > 1 && integer.charAt(i - 2) != '0'
                            || i > 2 && integer.charAt(i - 3) != '0') {
                        buffer.append(UNITS3[j / 4]);
                    }
                }
            } else {
                if (j % 4 == 0) {
                    // 插入[万]或者[亿]
                    buffer.append(UNITS3[j / 4]);
                }
                // 插入[拾]、[佰]或[仟]
                buffer.append(UNITS2[j % 4]);
                // 插入数字
                buffer.append(RMB_NUMS[n - '0']);
            }
        }
        return buffer.reverse().toString();
    }

    // 人民币数字
    private static final char[] RMB_NUMS = {'零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'};
    // 人民币单位
    private static final String[] UNITS2 = {"", "拾", "佰", "仟"};
    // 人民币大单位
    private static final String[] UNITS3 = {"", "万", "亿", "兆"};

    public static void main(String[] args) {
        String amount = "140,000.00";
        try {
            amount = amount.replaceAll(",", "");
            String chineseAmount = convertChineseAmount(amount);
            System.out.println("金额大写：" + chineseAmount);
        } catch (IllegalArgumentException e) {
            System.out.println(e.getMessage());
        }
    }
}