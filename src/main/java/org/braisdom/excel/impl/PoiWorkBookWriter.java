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
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.braisdom.excel.SheetWriter;
import org.braisdom.excel.WorkbookWriter;

import java.io.IOException;
import java.io.OutputStream;

public class PoiWorkBookWriter implements WorkbookWriter {

    private final HSSFWorkbook workbook;
    private final HSSFPalette palette;

    public PoiWorkBookWriter() {
        workbook = new HSSFWorkbook();
        palette = workbook.getCustomPalette();
    }

    @Override
    public void setActiveSheet(int sheetIndex) {
        workbook.setActiveSheet(sheetIndex);
    }

    @Override
    public SheetWriter createSheetWriter() {
        return new PoiSheetWriter(workbook, palette, workbook.createSheet());
    }

    @Override
    public void save(OutputStream outputStream) throws IOException {
        workbook.write(outputStream);
        workbook.close();
    }
}
