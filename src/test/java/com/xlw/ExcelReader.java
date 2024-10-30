package com.xlw;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        FileInputStream fileInputStream = new FileInputStream("/Users/xieliwei/Desktop/付款计划处理/付款单模板的副本.xlsx");
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = workbook.getSheetAt(2);

        for (int i = 0; i <= 5; i++) {
            System.out.println(i + ": " + sheet.getRow(i).getHeight());
        }

    }
}