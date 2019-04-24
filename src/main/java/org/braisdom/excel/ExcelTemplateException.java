package org.braisdom.excel;

public class ExcelTemplateException extends Exception {
    public ExcelTemplateException() {
    }

    public ExcelTemplateException(String message) {
        super(message);
    }

    public ExcelTemplateException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelTemplateException(Throwable cause) {
        super(cause);
    }

    public ExcelTemplateException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
