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
package org.braisdom.excel;

/**
 * Declaring the attributes of TemplateExcel template language.
 *
 * @author braisdom
 * @since 1.0.0
 */
public enum WorkbookAttribute {
    /** The name of sheet, which will be used for excel sheet name */
    NAME("name"),
    /** The height of row */
    HEIGHT("height"),
    /** The gaps is the distance rows above row  */
    ROW_GAPS("row-gaps"),
    /** The style of a row or a cell */
    STYLE("style"),
    /** The index of a cell */
    COLUMN("column"),
    /** Identifying a cell is quote prefixed */
    QUOTE_PREFIXED("quote-prefixed"),
    /** The column span of a cell */
    COLSPAN("colspan"),
    /** The row span of a cell */
    ROWSPAN("rowspan"),
    /** Identifying a cell width that will fit content */
    FIT_CONTENT("fit-content");

    private final String value;

    /**
     * Create a WorkbookAttribute with a string value.
     * @param value a string value of attribute
     */
    WorkbookAttribute(String value) {
        this.value = value;
    }

    /**
     * Returns a real value of attribute
     *
     * @return
     */
    public String getValue() {
        return this.value;
    }
}
