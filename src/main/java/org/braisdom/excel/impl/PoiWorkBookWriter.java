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
