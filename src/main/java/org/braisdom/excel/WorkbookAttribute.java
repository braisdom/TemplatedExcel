package org.braisdom.excel;

public enum WorkbookAttribute {
    NAME("name"),
    HEIGHT("height"),
    RPW_GAPS("row-gaps"),
    STYLE("style"),
    COLUMN("column"),
    QUOTE_PREFIXED("quote-prefixed"),
    COLSPAN("colspan"),
    ROWSPAN("rowspan"),
    FIT_CONTENT("fit-content");

    private final String value;

    WorkbookAttribute(String value) {
        this.value = value;
    }

    public String getValue() {
        return this.value;
    }
}
