package org.braisdom.excel;

public interface RowWriter extends StyleAware {

    public CellWriter createCellWriter(int cellIndex);

    public void setHeight(short height);
}
