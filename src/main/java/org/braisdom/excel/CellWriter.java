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
 * A writer of excel cell generating.
 *
 * @author braisdom
 * @since 1.0.0
 */
public interface CellWriter extends StyleAware {

    /**
     * Sets a text value of excel cell.
     * @param value the text value
     */
    public void setValue(String value);

    /**
     * Sets whether the cell is quote prefixed.
     * @param quotePrefix
     */
    public void setQuotePrefix(boolean quotePrefix);

    /**
     * Sets the merged rows and cells
     * @param rowIndex Index of first row
     * @param columnIndex  Index of first column
     * @param rowSpan  The distance for {@code rowIndex}
     * @param columnSpan The distance for {@code columnIndex}
     */
    public void setRowColumnSpan(int rowIndex, int columnIndex, int rowSpan, int columnSpan);

    /**
     * Sets whether cell width is fit content.
     * @param fitContent
     */
    public void setFitContent(boolean fitContent);
}
