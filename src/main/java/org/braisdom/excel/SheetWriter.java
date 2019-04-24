package org.braisdom.excel;

public interface SheetWriter {

    public void setName(String name);

    public RowWriter createRowWriter(int rowNum);

}
