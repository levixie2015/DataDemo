package com.xlw;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelCopyRange {

    public static void main(String[] args) throws IOException {
        String inputFilePath = "/Users/xieliwei/Desktop/付款计划处理/模版.xlsx";
        String outputFilePath = inputFilePath;

        FileInputStream fis = new FileInputStream(inputFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);

        // 定义要复制的区域（起始行，起始列，结束行，结束列）  
        int startRow = 0;
        int endRow = 6;
        int startCol = 0;
        int endCol = 6;

        int numRowSpace = 1;//数据间隔
        // 定义追加到的起始位置：第二行
        int appendStartRow = endRow + 1 + numRowSpace + 0 * 6;
        copyRangeWithStyles(sheet, startRow, startCol, endRow, endCol, appendStartRow, startCol);

        //定义追加到的起始位置：第三行
        int appendStartRow2 = appendStartRow + 1 + numRowSpace + 1 * 6;
        copyRangeWithStyles(sheet, appendStartRow - 1, startCol, appendStartRow2, endCol, appendStartRow2, startCol);


        FileOutputStream fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);
        fos.close();
        workbook.close();
        fis.close();
    }

    /**
     * 复制区域并带有样式
     *
     * @param sheet
     * @param startRow       起始行
     * @param startCol       起始列
     * @param endRow         结束行
     * @param endCol         结束列
     * @param appendStartRow
     * @param appendStartCol
     */
    private static void copyRangeWithStyles(XSSFSheet sheet, int startRow, int startCol, int endRow, int endCol, int appendStartRow, int appendStartCol) {
        Map<Integer, XSSFCellStyle> styleMap = new HashMap<>();

        for (int i = startRow; i <= endRow; i++) {
            XSSFRow sourceRow = sheet.getRow(i);

            XSSFRow destRow = sheet.createRow(appendStartRow + i - startRow);
            for (int j = startCol; j <= endCol; j++) {
                if (sourceRow != null) {
                    XSSFCell sourceCell = sourceRow.getCell(j);
                    XSSFCell destCell = destRow.createCell(appendStartCol + j - startCol);

                    if (sourceCell != null) {
                        copyCellValue(sourceCell, destCell);
                        copyCellStyle(sheet, sourceCell, destCell, styleMap);
                    }
                }
            }
        }

        // 复制合并单元格信息
        for (int k = 0; k < sheet.getNumMergedRegions(); k++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(k);
            if (mergedRegion.getFirstRow() >= startRow && mergedRegion.getLastRow() <= endRow &&
                    mergedRegion.getFirstColumn() >= startCol && mergedRegion.getLastColumn() <= endCol) {
                int newFirstRow = mergedRegion.getFirstRow() - startRow + appendStartRow;
                int newLastRow = mergedRegion.getLastRow() - startRow + appendStartRow;
                int newFirstCol = mergedRegion.getFirstColumn();
                int newLastCol = mergedRegion.getLastColumn();
                sheet.addMergedRegion(new CellRangeAddress(newFirstRow, newLastRow, newFirstCol, newLastCol));
            }
        }
    }

    private static void copyCellValue(XSSFCell sourceCell, XSSFCell destCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                destCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    destCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    destCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                destCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                destCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                destCell.setBlank();
                break;
            case ERROR:
                destCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
            default:
                break;
        }
    }

    private static void copyCellStyle(XSSFSheet sheet, XSSFCell sourceCell, XSSFCell destCell, Map<Integer, XSSFCellStyle> styleMap) {
        int styleIndex = sourceCell.getCellStyle().getIndex();
        XSSFCellStyle destStyle = styleMap.get(styleIndex);

        if (destStyle == null) {
            destStyle = sheet.getWorkbook().createCellStyle();
            XSSFCellStyle sourceStyle = sourceCell.getCellStyle();
            destStyle.cloneStyleFrom(sourceStyle);
            styleMap.put(styleIndex, destStyle);
        }
        destCell.setCellStyle(destStyle);
        destCell.getRow().setHeight(sourceCell.getRow().getHeight());
    }
}