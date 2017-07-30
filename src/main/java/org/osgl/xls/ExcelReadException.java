package org.osgl.xls;

import org.osgl.exception.UnexpectedException;

public class ExcelReadException extends UnexpectedException {
    public ExcelReadException(String message) {
        super(message);
    }

    public ExcelReadException(String message, Object... args) {
        super(message, args);
    }

    public ExcelReadException(Throwable cause) {
        super(cause);
    }

    public ExcelReadException(Throwable cause, String message, Object... args) {
        super(cause, message, args);
    }
}
