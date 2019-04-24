package org.braisdom.excel;

import java.io.IOException;
import java.io.OutputStream;

public interface WorkbookWriter {

    public void setActiveSheet(int sheetIndex);

    public SheetWriter createSheetWriter();

    public void save(OutputStream outputStream) throws IOException;
}
