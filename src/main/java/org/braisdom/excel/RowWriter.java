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
 * A writer of excel row generating.
 *
 * @author braisdom
 * @since 1.0.0
 */
public interface RowWriter extends StyleAware {

    /**
     * Sets the height of excel row.
     * @param height the height of row
     */
    public void setHeight(short height);

    /**
     * Creates a {@link CellWriter} for excel cell generating
     *
     * @param cellIndex the index of excel cell
     * @return a {@link CellWriter} for excel cell generating
     */
    public CellWriter createCellWriter(int cellIndex);
}
