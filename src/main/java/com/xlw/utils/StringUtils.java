package com.xlw.utils;

public class StringUtils {
    private StringUtils() {
    }

    public static String replaceAllBlank(String str) {
        if (str == null || str.equals("")) {
            return "";
        }
        return str.trim().replaceAll(" ", "");
    }

    public static void main(String[] args) {
        System.out.println(replaceAllBlank(" 22  22 "));
    }
}
