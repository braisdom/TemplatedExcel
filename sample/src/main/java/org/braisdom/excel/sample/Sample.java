package org.braisdom.excel.sample;

import org.braisdom.excel.WorkbookTemplate;
import org.braisdom.excel.impl.PoiWorkBookWriter;

import java.io.File;
import java.io.InputStreamReader;

public class Sample {

    public static void main(String[] args) throws Exception {
        InputStreamReader inputStreamReader = new InputStreamReader(Sample
                .class.getResourceAsStream("/template.xml"));
        File excelFile = new File("./sample.xls");
        WorkbookTemplate workbookTemplate = new WorkbookTemplate(inputStreamReader);
        workbookTemplate.process(new SampleTemplateDataSource(), new PoiWorkBookWriter(), excelFile);
    }
}
