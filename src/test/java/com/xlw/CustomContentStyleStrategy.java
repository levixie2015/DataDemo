package com.xlw;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.style.AbstractCellStyleStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

public class CustomContentStyleStrategy extends AbstractCellStyleStrategy {

//    @Override
//    protected WriteCellStyle contentCellStyle(CellData<?> cellData) {
//        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
//        contentWriteCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
//        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//        contentWriteCellStyle.setWrapped(true);
//
//        WriteFont contentWriteFont = new WriteFont();
//        contentWriteFont.setFontHeightInPoints((short) 11);
//        contentWriteFont.setFontName("宋体");
//        contentWriteCellStyle.setWriteFont(contentWriteFont);
//        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);
//        return contentWriteCellStyle;
//    }

    @Override
    protected void setContentCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
//        super.setContentCellStyle(cell, head, relativeRowIndex);

//        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
//        contentWriteCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
//        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//        contentWriteCellStyle.setWrapped(true);
//
//        WriteFont contentWriteFont = new WriteFont();
//        contentWriteFont.setFontHeightInPoints((short) 11);
//        contentWriteFont.setFontName("宋体");
//        contentWriteCellStyle.setWriteFont(contentWriteFont);
//        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);


        CellStyle cellStyle = cell.getCellStyle();
    }
}