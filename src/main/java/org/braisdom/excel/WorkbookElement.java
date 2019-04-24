package org.braisdom.excel;

public enum WorkbookElement {
    WORKBOOK("workbook"),
    SHEET("sheet"),
    ROW("row"),
    DATA_TABLE("data-table"),
    CELL("cell"),
    HEADER("header"),
    BODY("body");

    private String value;

    WorkbookElement(String value) {
        this.value = value;
    }

    public String getValue() {
        return this.value;
    }
}
