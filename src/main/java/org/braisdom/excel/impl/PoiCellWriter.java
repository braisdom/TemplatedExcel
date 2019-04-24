package org.braisdom.excel.impl;

import com.helger.css.decl.CSSDeclaration;
import com.helger.css.property.ECSSProperty;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.braisdom.excel.CellWriter;

import java.util.List;
import java.util.Map;

public class PoiCellWriter implements CellWriter {

    private final HSSFWorkbook hssfWorkbook;
    private final HSSFSheet hssfSheet;
    private final HSSFRow hssfRow;
    private final HSSFCell hssfCell;
    private final HSSFPalette palette;
    private final HSSFCellStyle hssfCellStyle;
    private final HSSFFont hssfFont;

    public PoiCellWriter(HSSFWorkbook hssfWorkbook,
                         HSSFPalette palette,
                         HSSFSheet hssfSheet,
                         HSSFRow hssfRow,
                         HSSFCell hssfCell) {
        this.hssfWorkbook = hssfWorkbook;
        this.palette = palette;
        this.hssfSheet = hssfSheet;
        this.hssfRow = hssfRow;
        this.hssfCell = hssfCell;

        this.hssfCellStyle = hssfWorkbook.createCellStyle();
        this.hssfFont = hssfWorkbook.createFont();

        this.hssfCell.setCellStyle(hssfCellStyle);
        this.hssfCellStyle.setFont(hssfFont);
    }

    @Override
    public void setValue(String value) {
        hssfCell.setCellValue(value);
    }

    @Override
    public void setQuotePrefix(boolean quotePrefix) {
        hssfCellStyle.setQuotePrefixed(quotePrefix);
    }

    @Override
    public void setFitContent(boolean fitContent) {
        if(fitContent) {
            int columnIndex = hssfCell.getColumnIndex();
            hssfSheet.autoSizeColumn(columnIndex, true);
        }
    }

    @Override
    public void setRowColumnSpan(int rowIndex, int columnIndex, int rowSpan, int columnSpan) {
        if(rowSpan >0 || columnSpan > 0) {
            CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex + rowSpan,
                    columnIndex, columnIndex + columnSpan);
            hssfSheet.addMergedRegionUnsafe(region);
        }
    }

    @Override
    public void setStyle(String style) {
        Map<String, List<CSSDeclaration>> cssDeclarations = StyleUtils.getStyle(style);
        if(!cssDeclarations.isEmpty()) {
            StyleUtils.setBackgroundColorStyle(palette, hssfCellStyle, cssDeclarations.get(ECSSProperty.BACKGROUND_COLOR.getName()));
            StyleUtils.setTextColor(palette, hssfFont, cssDeclarations.get(ECSSProperty.COLOR.getName()));

            StyleUtils.setHorizontalAlignment(hssfCellStyle, cssDeclarations.get(ECSSProperty.TEXT_ALIGN.getName()));
            StyleUtils.setVerticalAlignment(hssfCellStyle, cssDeclarations.get(ECSSProperty.VERTICAL_ALIGN.getName()));

            StyleUtils.setBorder(palette, hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER.getName()));
            StyleUtils.setLeftBorderStyle(hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER_BOTTOM_STYLE.getName()));
            StyleUtils.setRightBorderStyle(hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER_RIGHT_STYLE.getName()));
            StyleUtils.setTopBorderStyle(hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER_TOP_STYLE.getName()));
            StyleUtils.setBottomBorderStyle(hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER_BOTTOM_STYLE.getName()));
            StyleUtils.setLeftBorderColor(palette, hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER_LEFT_COLOR.getName()));
            StyleUtils.setRightBorderColor(palette, hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER_RIGHT_COLOR.getName()));
            StyleUtils.setTopBorderColor(palette, hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER_TOP_COLOR.getName()));
            StyleUtils.setBottomBorderColor(palette, hssfCellStyle, cssDeclarations.get(ECSSProperty.BORDER_BOTTOM_COLOR.getName()));

            StyleUtils.setWhiteSpaceStyle(hssfCellStyle, cssDeclarations.get(ECSSProperty.WHITE_SPACE.getName()));

            StyleUtils.setFontStyle(hssfFont, cssDeclarations.get(ECSSProperty.FONT.getName()));
            StyleUtils.setFontWeight(hssfFont, cssDeclarations.get(ECSSProperty.FONT_WEIGHT.getName()));
            StyleUtils.setFontFamily(hssfFont, cssDeclarations.get(ECSSProperty.FONT_FAMILY.getName()));
            StyleUtils.setFontSize(hssfFont, cssDeclarations.get(ECSSProperty.FONT_SIZE.getName()));
            StyleUtils.setFontStyle2(hssfFont, cssDeclarations.get(ECSSProperty.FONT_STYLE.getName()));
            StyleUtils.setTextDecoration(hssfFont, cssDeclarations.get(ECSSProperty.TEXT_DECORATION.getName()));
        }
    }
}
