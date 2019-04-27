/**
 * Copyright (C) 2019-2029 Braisdom Wang (www.joowing.com)
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

import java.io.IOException;
import java.io.OutputStream;

/**
 * A writer for excel workbook, which defines the basic capacities
 * for generating excel file.
 *
 * The workbook writer can be implemented by POI, etc.
 *
 * @author braisdom
 * @since 1.0.0
 */
public interface WorkbookWriter {

    /**
     * Make the assigned sheet index as active for the workbook.
     * Default index is 0.
     *
     * @param sheetIndex assigned sheet index will to be activated
     */
    public void setActiveSheet(int sheetIndex);

    /**
     * A factory method factory method for creating the sheet writer.
     *
     * @return a sheet writer who is used for creating excel sheet.
     */
    public SheetWriter createSheetWriter();

    /**
     * Save the workbook structure into a file or others.
     *
     * @param outputStream The file output stream or others.
     * @throws IOException If save the workbook structure failure.
     */
    public void save(OutputStream outputStream) throws IOException;
}
