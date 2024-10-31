package com.xlw;

import org.junit.Test;

import java.nio.file.Paths;

public class DemoTest {
    @Test
    public void testProcess() {
        String str = "/Users/xieliwei/Desktop/付款计划处理/模版.xlsx";
        System.out.println(Paths.get(str).getParent().toString());
    }
}