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
