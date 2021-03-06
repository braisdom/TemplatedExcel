/**
 * Copyright (C) 2019-2025 Braisdom Wang (www.joowing.com)
 * wangyonghe@msn.com
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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

    private boolean quotePrefix;

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
    }

    @Override
    public void setValue(String value) {
        hssfCell.setCellValue(value);
    }

    @Override
    public void setQuotePrefix(boolean quotePrefix) {
        this.quotePrefix = quotePrefix;
    }

    @Override
    public void setRowColumnSpan(int rowIndex, int columnIndex, int rowSpan, int columnSpan) {
        if (rowSpan > 0 || columnSpan > 0) {
            CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex + rowSpan,
                    columnIndex, columnIndex + columnSpan);
            hssfSheet.addMergedRegionUnsafe(region);
        }
    }

    @Override
    public void setStyle(String style) {
        Map<String, List<CSSDeclaration>> cssDeclarations = StyleUtils.getStyle(style);
        if (!cssDeclarations.isEmpty()) {
            HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
            HSSFFont hssfFont = hssfWorkbook.createFont();

            hssfCell.setCellStyle(hssfCellStyle);

            hssfCellStyle.setQuotePrefixed(quotePrefix);
            hssfCellStyle.setFont(hssfFont);

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
        } else if (quotePrefix) {
            HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
            hssfCell.setCellStyle(hssfCellStyle);

            hssfCellStyle.setQuotePrefixed(quotePrefix);
        }
    }
}
