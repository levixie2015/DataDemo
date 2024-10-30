package com.xlw;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.AbstractCellStyleStrategy;
import org.apache.poi.ss.usermodel.*;

public class CustomHeadStyleStrategy extends AbstractCellStyleStrategy {

//    @Override
//    protected WriteCellStyle headCellStyle(Head head) {
//        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
//        headWriteCellStyle.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
//        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
//        headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//        headWriteCellStyle.setWrapped(true);
//
//        WriteFont headWriteFont = new WriteFont();
//        headWriteFont.setFontHeightInPoints((short) 12);
//        headWriteFont.setFontName("微软雅黑");
//        headWriteFont.setBold(true);
//        headWriteCellStyle.setWriteFont(headWriteFont);
//        headWriteCellStyle.setBorderBottom(BorderStyle.THIN);
//        headWriteCellStyle.setBorderLeft(BorderStyle.THIN);
//        headWriteCellStyle.setBorderRight(BorderStyle.THIN);
//        headWriteCellStyle.setBorderTop(BorderStyle.THIN);
//        return headWriteCellStyle;
//    }


    @Override
    protected void setHeadCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
        super.setHeadCellStyle(cell, head, relativeRowIndex);
    }

    @Override
    protected void setContentCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
        super.setContentCellStyle(cell, head, relativeRowIndex);
    }
}