/**
 * Copyright (C) 2019-2025 Braisdom Wang (www.joowing.com)
 * wangyonghe@msn.com
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *         http://www.apache.org/licenses/LICENSE-2.0
 *
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
import org.braisdom.excel.CellWriter;
import org.braisdom.excel.RowWriter;

import java.util.List;
import java.util.Map;

public class PoiRowWriter implements RowWriter {

    private final HSSFWorkbook hssfWorkbook;
    private final HSSFSheet hssfSheet;
    private final HSSFRow hssfRow;
    private final HSSFPalette palette;
    private final HSSFCellStyle hssfCellStyle;
    private final HSSFFont hssfFont;

    public PoiRowWriter(HSSFWorkbook hssfWorkbook,
                        HSSFPalette palette,
                        HSSFSheet hssfSheet,
                        HSSFRow hssfRow) {
        this.hssfWorkbook = hssfWorkbook;
        this.palette = palette;
        this.hssfSheet = hssfSheet;
        this.hssfRow = hssfRow;
        this.hssfCellStyle = hssfWorkbook.createCellStyle();
        this.hssfFont = hssfWorkbook.createFont();

        this.hssfRow.setRowStyle(hssfCellStyle);
        this.hssfCellStyle.setFont(hssfFont);
    }

    @Override
    public void setHeight(short height) {
        hssfRow.setHeightInPoints(height);
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

    @Override
    public CellWriter createCellWriter(int cellColumn) {
        HSSFCell hssfCell = hssfRow.createCell(cellColumn);
        return new PoiCellWriter(hssfWorkbook, palette, hssfSheet, hssfRow, hssfCell);
    }
}
