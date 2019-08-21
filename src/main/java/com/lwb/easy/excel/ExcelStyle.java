package com.lwb.easy.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;

/**
 * excel样式设置工具类
 * @author liuweibo
 * @date 2019/8/20
 */
public interface ExcelStyle {

    /**
     * excel头部默认样式
     * @param book
     * @return
     */
    static CellStyle headerStyle(SXSSFWorkbook book) {
        CellStyle style = book.createCellStyle();

        // 基本样式
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 单元格水平方向样式
        style.setAlignment(HorizontalAlignment.CENTER);
        // 单元格垂直方向样式
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 自动换行
        style.setWrapText(true);

        // 设置字体
        Font font = book.createFont();
        font.setBold(true);
        style.setFont(font);

        return style;
    }

    /**
     * 设置合并单元格后的样式
     * @param addresses 合并的单元格坐标地址
     * @param sheet     所属sheet
     */
    static void setCellRangeAddress(List<CellRangeAddress> addresses, Sheet sheet) {
        addresses.forEach(address -> {
            RegionUtil.setBorderBottom(BorderStyle.THIN, address, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, address, sheet);
            RegionUtil.setBorderTop(BorderStyle.THIN, address, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, address, sheet);
        });
    }
}
