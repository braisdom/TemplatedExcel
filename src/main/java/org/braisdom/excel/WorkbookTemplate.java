/**
 * Copyright (C) 2019-2025 Braisdom Wang (www.joowing.com)
 * wangyonghe@msn.com
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.braisdom.excel;

import com.helger.css.propertyvalue.CCSSValue;
import org.apache.commons.io.IOUtils;
import org.braisdom.excel.impl.StyleUtils;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.IContext;

import java.io.*;
import java.util.*;

/**
 * This class is the base of the TemplatedExcel, it parses the raw template file
 * and generates the full structured data for writing the excel file.
 *
 * <p>
 * The <a href="https://github.com/thymeleaf/thymeleaf">thymeleaf</a> is used for the raw
 * template file parsing and generating the content which from Java objects. Structure the
 * template files with XML format, and the Dom4j will be used for XML content processing.
 * </p>
 * <p>
 * The <a href="https://github.com/phax/ph-css">ph-css</a> is used for parsing CSS text, and
 * making the CSS definition as Java accessibled.
 *
 * @author braisdom
 * @since 1.0.0
 */
public final class WorkbookTemplate {

    /**
     * The reader of the raw template content
     */
    private final Reader reader;

    /**
     * Create the WorkbookTemplate with string of template content
     */
    public WorkbookTemplate(String templateContent) {
        this(new StringReader(templateContent));
    }

    /**
     * Create the WorkbookTemplate with input reader
     */
    public WorkbookTemplate(Reader reader) {
        this.reader = reader;
    }

    /**
     * The main method for Excel file generating.
     *
     * @param templateDataSource the data source will be used for thymeleaf
     * @param workbookWriter     The Excel file generator.
     * @param excelFile          The Excel file reference.
     * @throws IOException       If read template source failure.
     * @throws DocumentException If parsing template content failure
     */
    public void process(TemplateDataSource templateDataSource,
                        WorkbookWriter workbookWriter,
                        File excelFile) throws IOException, DocumentException {
        String templateContent = processTemplate(templateDataSource);
        Document document = DocumentHelper.parseText(templateContent);
        List<Element> sheetElements = document.getRootElement().elements(WorkbookElement.SHEET.getValue());

        for (Element sheetElement : sheetElements) {
            SheetWriter sheetWriter = workbookWriter.createSheetWriter();
            List<Element> rowElements = sheetElement.elements();
            int realRowIndex = 0;

            processSheet(sheetElement, sheetWriter);

            for (int rowIndex = 0; rowIndex < rowElements.size(); rowIndex++) {
                Element rowElement = rowElements.get(rowIndex);
                int assignedRowGaps = StyleUtils.safeInteger(rowElement.attributeValue(WorkbookAttribute.ROW_GAPS.getValue()));
                // The gaps defined of row will be effecting the realRowIndex.
                realRowIndex += assignedRowGaps;
                if (rowElement.getName().equals(WorkbookElement.ROW.getValue())) {
                    processRowElement(realRowIndex + rowIndex, rowElement, sheetWriter);
                } else if (rowElement.getName().equals(WorkbookElement.DATA_TABLE.getValue())) {
                    // The DataTable has multiple rows, so the realRowIndex will be changed after it generated.
                    realRowIndex = processDataTableElement(realRowIndex, rowElement, sheetWriter);
                }
            }
        }
        workbookWriter.setActiveSheet(0);
        workbookWriter.save(new FileOutputStream(excelFile));
    }

    /**
     * Generating the row of Excel and attaching the style of row.
     *
     * @param rowIndex    The real row index of Excel.
     * @param rowElement  The structured row element who has the definition of CSS and the attributes.
     * @param sheetWriter The sheet writer of Excel, it will be used for creating row writer.
     */
    private void processRowElement(int rowIndex, Element rowElement, SheetWriter sheetWriter) {
        String style = rowElement.attributeValue(WorkbookAttribute.STYLE.getValue());
        String height = rowElement.attributeValue(WorkbookAttribute.HEIGHT.getValue());
        List<Element> cellElements = rowElement.elements(WorkbookElement.CELL.getValue());

        RowWriter rowWriter = sheetWriter.createRowWriter(rowIndex);

        if (style != null)
            rowWriter.setStyle(style);

        if (height != null)
            rowWriter.setHeight(Short.parseShort(height));

        for (int cellIndex = 0; cellIndex < cellElements.size(); cellIndex++) {
            Element cellElement = cellElements.get(cellIndex);
            processCellElement(rowIndex, style, cellIndex, cellElement, rowWriter);
        }
    }

    /**
     * Generating the cell of Excel and attaching the style of cell.
     *
     * @param rowIndex    The real row index of Excel.
     * @param rowStyle    The row style for cell inheriting.
     * @param columnIndex The real column index of Excel.
     * @param cellElement The structured cell element who has the definition of CSS and the attributes.
     * @param rowWriter   The row writer of Excel, it will be used for creating cell writer.
     */
    private void processCellElement(int rowIndex, String rowStyle, int columnIndex, Element cellElement, RowWriter rowWriter) {
        String rawAssignedCellIndex = cellElement.attributeValue(WorkbookAttribute.COLUMN.getValue());
        String rawRowSpan = cellElement.attributeValue(WorkbookAttribute.ROWSPAN.getValue());
        String rawColumnSpan = cellElement.attributeValue(WorkbookAttribute.COLSPAN.getValue());
        String style = cellElement.attributeValue(WorkbookAttribute.STYLE.getValue());

        boolean quotePrefix = Boolean.valueOf(cellElement.attributeValue(WorkbookAttribute.QUOTE_PREFIXED.getValue()));

        int assignedColumnIndex = rawAssignedCellIndex == null ? columnIndex : StyleUtils.safeInteger(rawAssignedCellIndex);
        int rowSpan = StyleUtils.safeInteger(rawRowSpan);
        int columnSpan = StyleUtils.safeInteger(rawColumnSpan);

        CellWriter cellWriter = rowWriter.createCellWriter(assignedColumnIndex);
        cellWriter.setValue(cellElement.getText().trim());
        cellWriter.setQuotePrefix(quotePrefix);
        cellWriter.setRowColumnSpan(rowIndex, assignedColumnIndex, rowSpan, columnSpan);

        if (CCSSValue.INHERIT.equals(style) && rowStyle != null)
            cellWriter.setStyle(rowStyle);
        if (style != null)
            cellWriter.setStyle(style);
    }

    /**
     * Parsing the data-table tags and generating the rows and cells of Excel
     *
     * @param rowIndex         current row index of Excel
     * @param dataTableElement the data table definition
     * @param sheetWriter      the sheet writer of Excel
     * @return
     */
    private int processDataTableElement(int rowIndex, Element dataTableElement, SheetWriter sheetWriter) {
        Element headerElement = dataTableElement.element(WorkbookElement.HEADER.getValue());
        Element bodyElement = dataTableElement.element(WorkbookElement.BODY.getValue());

        Map<Element, Boolean> fitColumns = new HashMap<>();

        int dataTableRowIndex = rowIndex;
        int headerRowCount = 0;
        int columnCount = -1;
        boolean fitContent = false;

        if (headerElement != null) {
            List<Element> rowElements = headerElement.elements(WorkbookElement.ROW.getValue());
            for (int headerRowIndex = 0; headerRowIndex < rowElements.size(); headerRowIndex++) {
                Element rowElement = rowElements.get(headerRowIndex);
                int assignedRowGaps = StyleUtils.safeInteger(rowElement.attributeValue(WorkbookAttribute.ROW_GAPS.getValue()));
                columnCount = rowElement.elements(WorkbookElement.CELL.getValue()).size();
                fitContent = Boolean.valueOf(rowElement.attributeValue(WorkbookAttribute.FIT_CONTENT.getValue()));

                dataTableRowIndex += assignedRowGaps;
                processRowElement(dataTableRowIndex + headerRowCount, rowElement, sheetWriter);
                headerRowCount++;
            }
        }

        if (bodyElement != null) {
            List<Element> rowElements = bodyElement.elements(WorkbookElement.ROW.getValue());

            for (int bodyRowIndex = 0; bodyRowIndex < rowElements.size(); bodyRowIndex++) {
                Element rowElement = rowElements.get(bodyRowIndex);
                int assignedRowGaps = StyleUtils.safeInteger(rowElement.attributeValue(WorkbookAttribute.ROW_GAPS.getValue()));
                dataTableRowIndex += assignedRowGaps;
                processRowElement(dataTableRowIndex + headerRowCount + bodyRowIndex, rowElement, sheetWriter);
            }
        }

        // Resize the column width finally because the POI performance problem.
        if (fitContent) {
            for (int i = 0; i < columnCount; i++) {
                sheetWriter.autoSizeColumn(i);
            }
        }

        return dataTableRowIndex;
    }

    private void processSheet(Element sheetElement, SheetWriter sheetWriter) {
        sheetWriter.setName(sheetElement.attributeValue(WorkbookAttribute.NAME.getValue()));
    }

    private String processTemplate(TemplateDataSource templateDataSource) throws IOException {
        String rawTemplateContent = IOUtils.toString(reader);
        TemplateEngine templateEngine = new TemplateEngine();
        IContext context = new DefaultThymeleafContext(templateDataSource);
        return templateEngine.process(rawTemplateContent, context);
    }

    private class DefaultThymeleafContext implements IContext {

        private final TemplateDataSource templateDataSource;

        public DefaultThymeleafContext(TemplateDataSource templateDataSource) {
            this.templateDataSource = templateDataSource;
        }

        @Override
        public Locale getLocale() {
            return Locale.getDefault();
        }

        @Override
        public boolean containsVariable(String name) {
            return this.templateDataSource.getDataNames().contains(name);
        }

        @Override
        public Set<String> getVariableNames() {
            return this.templateDataSource.getDataNames();
        }

        @Override
        public Object getVariable(String name) {
            return this.templateDataSource.getData(name);
        }
    }

}
