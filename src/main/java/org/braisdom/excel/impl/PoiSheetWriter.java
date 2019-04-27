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

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.braisdom.excel.RowWriter;
import org.braisdom.excel.SheetWriter;

public class PoiSheetWriter implements SheetWriter {

    private final HSSFSheet hssfSheet;
    private final HSSFWorkbook hssfWorkbook;
    private final HSSFPalette palette;

    public PoiSheetWriter(HSSFWorkbook hssfWorkbook, HSSFPalette palette, HSSFSheet hssfSheet) {
        this.hssfWorkbook = hssfWorkbook;
        this.palette = palette;
        this.hssfSheet = hssfSheet;
    }

    @Override
    public void setName(String name) {
        hssfWorkbook.setSheetName(hssfWorkbook.getSheetIndex(hssfSheet), name);
    }

    @Override
    public RowWriter createRowWriter(int rowNum) {
        HSSFRow hssfRow = hssfSheet.createRow(rowNum);
        return new PoiRowWriter(hssfWorkbook, palette, hssfSheet, hssfRow);
    }
}
