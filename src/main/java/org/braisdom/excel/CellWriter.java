package org.braisdom.excel;

public interface CellWriter extends StyleAware {

    public void setValue(String value);

    public void setQuotePrefix(boolean quotePrefix);

    public void setRowColumnSpan(int rowIndex, int columnIndex, int rowSpan, int columnSpan);

    public void setFitContent(boolean fitContent);
}
