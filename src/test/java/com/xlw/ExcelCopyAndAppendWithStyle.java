package com.xlw;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelCopyAndAppendWithStyle {

    public static void main(String[] args) {
        String excelFilePath = "/Users/xieliwei/Desktop/付款计划处理/模版.xlsx";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // 获取第一个工作表

            // 指定要复制的区域（例如：A1:C3）
            int startRow = 0;
            int endRow = 6;
            int startCol = 0;
            int endCol = 6;

            // 获取目标开始行（追加的起始位置）
            int appendStartRow = endRow + 1 + 1; // 加1是为了在复制区域下方留出空行（可选）

            // 复制指定区域并保留样式
            for (int row = startRow; row <= endRow; row++) {
                Row sourceRow = sheet.getRow(row);
                Row targetRow = sheet.createRow(appendStartRow + row - startRow);
                copyRowStyle(sourceRow, targetRow);

                if (sourceRow != null) {
                    for (int col = startCol; col <= endCol; col++) {
                        Cell sourceCell = sourceRow.getCell(col);
                        Cell targetCell = targetRow.createCell(col);

                        if (sourceCell != null) {
                            copyCellValue(sourceCell, targetCell);
                            copyCellStyle(sourceCell, targetCell, workbook);
                        }
                    }
                }
            }

            // 写入到文件
            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                workbook.write(fos);
            }

            System.out.println("指定区域已成功复制并追加到工作表中，且样式已保留。");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                targetCell.setBlank();
                break;
            default:
                break;
        }
    }

    private static void copyRowStyle(Row sourceRow, Row targetRow) {
        CellStyle sourceRowStyle = sourceRow.getRowStyle();
        targetRow.setRowStyle(sourceRowStyle);
        targetRow.setHeight(sourceRow.getHeight());
        targetRow.setHeightInPoints(sourceRow.getHeightInPoints());
    }

    private static void copyCellStyle(Cell sourceCell, Cell targetCell, Workbook workbook) {
        CellStyle sourceCellStyle = sourceCell.getCellStyle();
//        targetCell.setCellStyle(sourceCell.getCellStyle());
        XSSFCellStyle targetCellStyle = (XSSFCellStyle) workbook.createCellStyle();
//
//        // 复制样式属性（这里只复制了一些常见的样式属性，你可以根据需要添加更多）
        targetCellStyle.setAlignment(sourceCellStyle.getAlignment());
        targetCellStyle.setVerticalAlignment(sourceCellStyle.getVerticalAlignment());
        targetCellStyle.setBorderBottom(sourceCellStyle.getBorderBottom());
        targetCellStyle.setBorderLeft(sourceCellStyle.getBorderLeft());
        targetCellStyle.setBorderRight(sourceCellStyle.getBorderRight());
        targetCellStyle.setBorderTop(sourceCellStyle.getBorderTop());
        targetCellStyle.setBottomBorderColor(sourceCellStyle.getBottomBorderColor());
        targetCellStyle.setLeftBorderColor(sourceCellStyle.getLeftBorderColor());
        targetCellStyle.setRightBorderColor(sourceCellStyle.getRightBorderColor());
        targetCellStyle.setTopBorderColor(sourceCellStyle.getTopBorderColor());
        targetCellStyle.setFillForegroundColor(sourceCellStyle.getFillForegroundColor());
        targetCellStyle.setFillPattern(sourceCellStyle.getFillPattern());
        targetCellStyle.setFont(workbook.getFontAt(sourceCellStyle.getFontIndexAsInt()));
        targetCellStyle.setWrapText(sourceCellStyle.getWrapText());

//        // 注意：对于更复杂的样式（如数据格式、旋转角度等），你可能需要额外的处理
//
        targetCell.setCellStyle(targetCellStyle);
    }
}