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
 * A writer of excel sheet generating.
 *
 * @author braisdom
 * @since 1.0.0
 */
public interface SheetWriter {

    /**
     * Sets the name of excel sheet.
     * @param name assigned name
     */
    public void setName(String name);

    /**
     * Create a {@link CellWriter} for excel generating.
     *
     * @param rowNum assigned row index
     * @return a {@link RowWriter} of excel row generating.
     */
    public RowWriter createRowWriter(int rowNum);

    /**
     * Make the column fit to the content.
     *
     * @since 1.0.3
     */
    public void autoSizeColumn(int columnIndex);

}
