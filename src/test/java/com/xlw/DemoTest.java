package com.xlw;

import org.junit.Test;

import java.nio.file.Paths;

public class DemoTest {
    @Test
    public void testSay() {
        String str = "/Users/xieliwei/Desktop/付款计划处理/付款单模板.xlsx";
        System.out.println(Paths.get(str).getParent().toString());
    }
}