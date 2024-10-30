package com.xlw.utils;

import com.xlw.model.PayPlanTemplateData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class PayPlanExcelUtils {
    public static void write(List<PayPlanTemplateData> dataList, String sheetName, String fileName) {
        // 创建一个新的工作簿
        Workbook workbook = new XSSFWorkbook();
        // 创建一个新的工作表
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.setColumnWidth(0, 3776); // 0表示第一列的索引,// Excel列宽是以字符宽度为单位的，一个字符宽度等于256个单位
        sheet.setColumnWidth(1, 2752); // 0表示第一列的索引,// Excel列宽是以字符宽度为单位的，一个字符宽度等于256个单位
        sheet.setColumnWidth(2, 1376); // 0表示第一列的索引,// Excel列宽是以字符宽度为单位的，一个字符宽度等于256个单位
        sheet.setColumnWidth(3, 4704); // 0表示第一列的索引,// Excel列宽是以字符宽度为单位的，一个字符宽度等于256个单位
        sheet.setColumnWidth(4, 4096); // 0表示第一列的索引,// Excel列宽是以字符宽度为单位的，一个字符宽度等于256个单位
        sheet.setColumnWidth(5, 4032); // 0表示第一列的索引,// Excel列宽是以字符宽度为单位的，一个字符宽度等于256个单位

        int currentRow = 0;
        //遍历数据
        for (int i = 0; i < dataList.size(); i++) {
            PayPlanTemplateData data = dataList.get(i);
            setFirstRow(sheet, currentRow, data);
            setSecondRow(workbook, sheet, currentRow, data);
            setThirdRow(workbook, sheet, currentRow, data);
            setFourthRow(workbook, sheet, currentRow, data);
            setFifthRow(workbook, sheet, currentRow, data);
            setSixthRow(workbook, sheet, currentRow, data);
            setSeventhRow(workbook, sheet, currentRow, data);
            setEighthRow(workbook, sheet, currentRow, data);
            currentRow = setNinthRow(workbook, sheet, currentRow, data);
            if (i > 1 && i % 3 == 2) {
                currentRow = currentRow + 8;
            } else {
                currentRow = currentRow + 5;
            }
        }

        // 将工作簿写入文件
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 第一行数据
     *
     * @param sheet
     */
    private static void setFirstRow(Sheet sheet, int currentData, PayPlanTemplateData data) {
        int currentRorw = currentData + 0;
        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        setRowCell(sheet, currentRorw, currentRorw, 0, 6, null, "", false);
    }

    /**
     * 第二行数据
     *
     * @param sheet
     */
    private static void setSecondRow(Workbook workbook, Sheet sheet, int currentData, PayPlanTemplateData data) {
        int currentRorw = currentData + 1;
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        cellStyle.setFont(font);

        //创建一个字体对象并设置其属性
        font.setFontName("微软雅黑"); // 设置字体名称
        font.setFontHeightInPoints((short) 16); // 设置字体大小
        font.setItalic(false); // 斜体
        cellStyle.setFont(font);
        // 设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//HorizontalAlignment.CENTER
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//VerticalAlignment.CENTER
        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        Cell cell = setRowCell(sheet, currentRorw, currentRorw, 0, 6, cellStyle, data.getTitile(), false);

        // 将单元格样式应用到目标单元格上
        cell.setCellStyle(cellStyle);
    }

    /**
     * 第三行数据
     *
     * @param sheet
     */
    private static void setThirdRow(Workbook workbook, Sheet sheet, int currentData, PayPlanTemplateData data) {
        int currentRorw = currentData + 2;
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        cellStyle.setFont(font);
        //创建一个字体对象并设置其属性
        font.setFontName("等线"); // 设置字体名称
        font.setFontHeightInPoints((short) 12); // 设置字体大小
        cellStyle.setFont(font);
        // 设置水平靠右
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);//HorizontalAlignment.RIGHT
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//VerticalAlignment.CENTER

        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        Cell cell = setRowCell(sheet, currentRorw, currentRorw, 0, 6, cellStyle, "编号：" + data.getSerialNo(), false);
        // 将单元格样式应用到目标单元格上
        cell.setCellStyle(cellStyle);
    }

    /**
     * 第四行数据
     *
     * @param sheet
     */
    private static void setFourthRow(Workbook workbook, Sheet sheet, int currentData, PayPlanTemplateData data) {
        int currentRorw = currentData + 3;
        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        setRowCell(sheet, currentRorw, currentRorw, 0, 6, null, "", false);
    }

    /**
     * 第五行数据
     *
     * @param sheet
     */
    private static void setFifthRow(Workbook workbook, Sheet sheet, int currentData, PayPlanTemplateData data) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        cellStyle.setFont(font);
        //创建一个字体对象并设置其属性
        font.setFontName("等线"); // 设置字体名称
        font.setFontHeightInPoints((short) 12); // 设置字体大小
        cellStyle.setFont(font);
        // 设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//HorizontalAlignment.CENTER
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//VerticalAlignment.CENTER
        cellStyle.setBorderTop(BorderStyle.THIN);//BorderStyle.THIN
        cellStyle.setBorderLeft(BorderStyle.THIN);//BorderStyle.THIN
        cellStyle.setBorderRight(BorderStyle.THIN);//BorderStyle.THIN
        cellStyle.setBorderBottom(BorderStyle.THIN);//BorderStyle.THIN

        int currentRorw = currentData + 4;
        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        Cell cell1 = setRowCell(sheet, currentRorw, currentRorw, 0, 0, cellStyle, "供应商全称", true);
        cell1.setCellStyle(cellStyle);

        Cell cell2 = setRowCell(sheet, currentRorw, currentRorw, 1, 3, cellStyle, data.getSupplierName(), true);
        cell2.setCellStyle(cellStyle);


        Cell cell3 = setRowCell(sheet, currentRorw, currentRorw, 4, 4, cellStyle, "开户行", true);
        cell3.setCellStyle(cellStyle);

        Cell cell4 = setRowCell(sheet, currentRorw, currentRorw, 5, 6, cellStyle, data.getOpeningBank(), true);
        cell4.setCellStyle(cellStyle);
    }

    /**
     * 第六行数据
     *
     * @param sheet
     */
    private static void setSixthRow(Workbook workbook, Sheet sheet, int currentData, PayPlanTemplateData data) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        cellStyle.setFont(font);
        //创建一个字体对象并设置其属性
        font.setFontName("等线"); // 设置字体名称
        font.setFontHeightInPoints((short) 12); // 设置字体大小
        cellStyle.setFont(font);
        // 设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//HorizontalAlignment.CENTER
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//VerticalAlignment.CENTER


        int currentRorw = currentData + 5;
        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        Cell cell1 = setRowCell(sheet, currentRorw, currentRorw, 0, 0, cellStyle, "供应商编码", true);
        cell1.setCellStyle(cellStyle);

        Cell cell2 = setRowCell(sheet, currentRorw, currentRorw, 1, 3, cellStyle, data.getSupplierCode(), true);
        cell2.setCellStyle(cellStyle);

        Cell cell3 = setRowCell(sheet, currentRorw, currentRorw, 4, 4, cellStyle, "账号", true);
        cell3.setCellStyle(cellStyle);

        Cell cell4 = setRowCell(sheet, currentRorw, currentRorw, 5, 6, cellStyle, data.getAccountNo(), true);
        cell4.setCellStyle(cellStyle);
    }

    /**
     * 第七行数据
     *
     * @param sheet
     */
    private static void setSeventhRow(Workbook workbook, Sheet sheet, int currentData, PayPlanTemplateData data) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        cellStyle.setFont(font);
        //创建一个字体对象并设置其属性
        font.setFontName("等线"); // 设置字体名称
        font.setFontHeightInPoints((short) 12); // 设置字体大小
        cellStyle.setFont(font);
        // 设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//HorizontalAlignment.CENTER
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//VerticalAlignment.CENTER


        int currentRorw = currentData + 6;
        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        Cell cell1 = setRowCell(sheet, currentRorw, currentRorw, 0, 0, cellStyle, "付款总额", true);
        cell1.setCellStyle(cellStyle);

        Cell cell2 = setRowCell(sheet, currentRorw, currentRorw, 1, 3, cellStyle, data.getPayAmount(), true);
        cell2.setCellStyle(cellStyle);

        Cell cell3 = setRowCell(sheet, currentRorw, currentRorw, 4, 4, cellStyle, "行号", true);
        cell3.setCellStyle(cellStyle);

        Cell cell4 = setRowCell(sheet, currentRorw, currentRorw, 5, 6, cellStyle, data.getOpeningBankNo(), true);
        cell4.setCellStyle(cellStyle);
    }

    /**
     * 第八行数据
     *
     * @param sheet
     */
    private static void setEighthRow(Workbook workbook, Sheet sheet, int currentData, PayPlanTemplateData data) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        cellStyle.setFont(font);
        //创建一个字体对象并设置其属性
        font.setFontName("等线"); // 设置字体名称
        font.setFontHeightInPoints((short) 12); // 设置字体大小
        cellStyle.setFont(font);
        // 设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//HorizontalAlignment.CENTER
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//VerticalAlignment.CENTER

        int currentRorw = currentData + 7;
        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        Cell cell1 = setRowCell(sheet, currentRorw, currentRorw, 0, 0, cellStyle, "付款方式", true);
        cell1.setCellStyle(cellStyle);

        Cell cell2 = setRowCell(sheet, currentRorw, currentRorw, 1, 2, cellStyle, data.getPaymentMethod(), true);
        cell2.setCellStyle(cellStyle);

        Cell cell3 = setRowCell(sheet, currentRorw, currentRorw, 3, 3, cellStyle, data.getAmount(), true);
        cell3.setCellStyle(cellStyle);

        Cell cell4 = setRowCell(sheet, currentRorw, currentRorw, 4, 4, cellStyle, "付款单位", true);
        cell4.setCellStyle(cellStyle);

        Cell cell5 = setRowCell(sheet, currentRorw, currentRorw, 5, 6, cellStyle, data.getPayer(), true);
        cell5.setCellStyle(cellStyle);
    }

    /**
     * 第九行数据
     *
     * @param sheet
     */
    private static int setNinthRow(Workbook workbook, Sheet sheet, int currentData, PayPlanTemplateData data) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        cellStyle.setFont(font);
        //创建一个字体对象并设置其属性
        font.setFontName("等线"); // 设置字体名称
        font.setFontHeightInPoints((short) 12); // 设置字体大小
        cellStyle.setFont(font);
        // 设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//HorizontalAlignment.CENTER
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//VerticalAlignment.CENTER


        int currentRorw = currentData + 8;
        // 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
        Cell cell1 = setRowCell(sheet, currentRorw, currentRorw, 0, 0, cellStyle, "经办人", true);
        cell1.setCellStyle(cellStyle);

        Cell cell2 = setRowCell(sheet, currentRorw, currentRorw, 1, 3, cellStyle, data.getPaymentMethod(), true);
        cell2.setCellStyle(cellStyle);

        Cell cell3 = setRowCell(sheet, currentRorw, currentRorw, 4, 4, cellStyle, "财务", true);
        cell3.setCellStyle(cellStyle);

        Cell cell4 = setRowCell(sheet, currentRorw, currentRorw, 5, 6, cellStyle, data.getPayer(), true);
        cell4.setCellStyle(cellStyle);
        return currentRorw;
    }

    /**
     * 定义要合并的单元格范围（起始行，结束行，起始列，结束列）
     *
     * @param sheet
     * @param firstRow 起始行
     * @param lastRow  结束行
     * @param firstCol 起始列
     * @param lastCol  结束列
     * @param value    单元格值
     */
    private static Cell setRowCell(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, CellStyle cellStyle, String value, boolean setBorder) {
        if (firstCol != lastCol) {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            // 合并单元格
            sheet.addMergedRegion(cellRangeAddress);
        }
        // 创建一个单元格（这个单元格实际上是在合并区域的左上角）
        Row row = sheet.getRow(firstRow);
        if (row == null) {
            row = sheet.createRow(firstRow);
        }
        Cell cell = row.getCell(firstCol);
        if (cell == null) {
            cell = row.createCell(firstCol);
        }

        if (setBorder) {
            setRegionStyle(sheet, cellStyle, firstRow, lastRow, firstCol, lastCol);
        }
        cell.setCellValue(value);
        return cell;
    }

    /**
     * 设置边框，并解决跨行边框失效问题
     *
     * @param sheet
     * @param cellStyle
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    private static void setRegionStyle(Sheet sheet, CellStyle cellStyle, int firstRow, int lastRow, int firstCol, int lastCol) {
        for (int i = firstRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            //循环设置单元格样式
            Cell cell = null;
            for (int j = firstCol; j <= lastCol; j++) {
                cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                }
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cell.setCellStyle(cellStyle);
            }
        }
    }
}
