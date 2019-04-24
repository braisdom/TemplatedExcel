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
import java.util.List;
import java.util.Locale;
import java.util.Set;

public final class WorkbookTemplate {

    private final Reader reader;

    public WorkbookTemplate(String templateContent) {
        this(new StringReader(templateContent));
    }

    public WorkbookTemplate(Reader reader) {
        this.reader = reader;
    }

    public void process(TemplateDataSource templateDataSource,
                        WorkbookWriter workbookWriter, File excelFile) throws IOException, DocumentException {
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
                int assignedRowGaps = StyleUtils.safeInteger(rowElement.attributeValue(WorkbookAttribute.RPW_GAPS.getValue()));
                realRowIndex += (assignedRowGaps + rowIndex);
                if (rowElement.getName().equals(WorkbookElement.ROW.getValue())) {
                    processRowElement(realRowIndex, rowElement, sheetWriter);
                } else if (rowElement.getName().equals(WorkbookElement.DATA_TABLE.getValue())) {
                    realRowIndex = processDataTableElement(realRowIndex, rowElement, sheetWriter);
                }
            }
        }
        workbookWriter.setActiveSheet(0);
        workbookWriter.save(new FileOutputStream(excelFile));
    }

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

    private void processCellElement(int rowIndex, String rowStyle, int columnIndex, Element cellElement, RowWriter rowWriter) {
        String rawAssignedCellIndex = cellElement.attributeValue(WorkbookAttribute.COLUMN.getValue());
        String rawRowSpan = cellElement.attributeValue(WorkbookAttribute.ROWSPAN.getValue());
        String rawColumnSpan = cellElement.attributeValue(WorkbookAttribute.COLSPAN.getValue());
        String style = cellElement.attributeValue(WorkbookAttribute.STYLE.getValue());

        boolean quotePrefix = Boolean.valueOf(cellElement.attributeValue(WorkbookAttribute.QUOTE_PREFIXED.getValue()));
        boolean fitContent = Boolean.valueOf(cellElement.attributeValue(WorkbookAttribute.FIT_CONTENT.getValue()));
        int assignedColumnIndex = rawAssignedCellIndex == null ? columnIndex : StyleUtils.safeInteger(rawAssignedCellIndex);
        int rowSpan = StyleUtils.safeInteger(rawRowSpan);
        int columnSpan = StyleUtils.safeInteger(rawColumnSpan);

        CellWriter cellWriter = rowWriter.createCellWriter(assignedColumnIndex);
        cellWriter.setValue(cellElement.getText());
        cellWriter.setQuotePrefix(quotePrefix);
        cellWriter.setFitContent(fitContent);
        cellWriter.setRowColumnSpan(rowIndex, assignedColumnIndex, rowSpan, columnSpan);

        if (CCSSValue.INHERIT.equals(style) && rowStyle != null)
            cellWriter.setStyle(rowStyle);
        if (style != null)
            cellWriter.setStyle(style);
    }

    private int processDataTableElement(int rowIndex, Element dataTableElement, SheetWriter sheetWriter) {
        Element headerElement = dataTableElement.element(WorkbookElement.HEADER.getValue());
        Element bodyElement = dataTableElement.element(WorkbookElement.BODY.getValue());
        int dataTableRowIndex = rowIndex;
        int headerRowCount = 0;

        if (headerElement != null) {
            List<Element> rowElements = headerElement.elements(WorkbookElement.ROW.getValue());
            for (int headerRowIndex = 0; headerRowIndex < rowElements.size(); headerRowIndex++) {
                Element rowElement = rowElements.get(0);
                int assignedRowGaps = StyleUtils.safeInteger(rowElement.attributeValue(WorkbookAttribute.RPW_GAPS.getValue()));
                dataTableRowIndex += (assignedRowGaps + headerRowIndex);
                processRowElement(dataTableRowIndex, rowElement, sheetWriter);
                headerRowCount++;
            }
        }

        if (bodyElement != null) {
            List<Element> rowElements = bodyElement.elements(WorkbookElement.ROW.getValue());
            for (int bodyRowIndex = 0; bodyRowIndex < rowElements.size(); bodyRowIndex++) {
                Element rowElement = rowElements.get(0);
                int assignedRowGaps = StyleUtils.safeInteger(rowElement.attributeValue(WorkbookAttribute.RPW_GAPS.getValue()));
                dataTableRowIndex += (assignedRowGaps + bodyRowIndex);
                processRowElement(dataTableRowIndex + headerRowCount, rowElement, sheetWriter);
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
