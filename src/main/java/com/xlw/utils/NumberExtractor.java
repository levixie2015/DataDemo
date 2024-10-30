
package com.xlw.utils;

import java.util.ArrayList;
import java.util.List;  
import java.util.regex.Matcher;  
import java.util.regex.Pattern;  
  
public class NumberExtractor {  
  
    // 正则表达式模式，用于匹配整数和小数  
    private static final Pattern NUMBER_PATTERN = Pattern.compile("-?\\d+(\\.\\d+)?");  
  
    /**  
     * 从字符串中提取所有数字（包括整数和小数），并返回它们的字符串列表。  
     *  
     * @param input 输入字符串  
     * @return 数字字符串列表  
     */  
    public static List<String> extractNumbers(String input) {  
        List<String> numbers = new ArrayList<>();  
        Matcher matcher = NUMBER_PATTERN.matcher(input);  
  
        while (matcher.find()) {  
            numbers.add(matcher.group());  
        }  
  
        return numbers;  
    }  
  
    /**  
     * 从字符串中提取所有数字（包括整数和小数），并返回它们的双精度列表。  
     *  
     * @param input 输入字符串  
     * @return 数字双精度列表  
     */  
    public static List<Double> extractNumbersAsDoubles(String input) {  
        List<Double> numbers = new ArrayList<>();  
        Matcher matcher = NUMBER_PATTERN.matcher(input);  
  
        while (matcher.find()) {  
            numbers.add(Double.parseDouble(matcher.group()));  
        }  
  
        return numbers;  
    }  
  
    /**  
     * 从字符串中提取所有整数，并返回它们的字符串列表。  
     *  
     * @param input 输入字符串  
     * @return 整数字符串列表  
     */  
    public static List<String> extractIntegers(String input) {  
        List<String> integers = new ArrayList<>();  
        Matcher matcher = Pattern.compile("-?\\d+").matcher(input);  
  
        while (matcher.find()) {  
            integers.add(matcher.group());  
        }  
  
        return integers;  
    }  
  
    /**  
     * 从字符串中提取所有整数，并返回它们的整数列表。  
     *  
     * @param input 输入字符串  
     * @return 整数列表  
     */  
    public static List<Integer> extractIntegersAsInts(String input) {  
        List<Integer> integers = new ArrayList<>();  
        Matcher matcher = Pattern.compile("-?\\d+").matcher(input);  
  
        while (matcher.find()) {  
            integers.add(Integer.parseInt(matcher.group()));  
        }  
  
        return integers;  
    }  
  
    public static void main(String[] args) {  
//        String testString = "The price is 123.45 dollars and the discount is -20% or 20.5 dollars.";
        String testString = "采购9月付款计划";

        // 测试提取所有数字（包括整数和小数）  
        List<String> numbers = extractNumbers(testString);  
        System.out.println("Extracted Numbers (String): " + numbers);  
  
        // 测试提取所有数字（包括整数和小数）作为双精度列表  
        List<Double> numbersAsDoubles = extractNumbersAsDoubles(testString);  
        System.out.println("Extracted Numbers (Double): " + numbersAsDoubles);  
  
        // 测试提取所有整数  
        List<String> integers = extractIntegers(testString);  
        System.out.println("Extracted Integers (String): " + integers);  
  
        // 测试提取所有整数作为整数列表  
        List<Integer> integersAsInts = extractIntegersAsInts(testString);  
        System.out.println("Extracted Integers (Int): " + integersAsInts);  
    }  
}