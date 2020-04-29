package com.kyle.excel.exception;

/**
 * @package: com.kyle.excel.exception
 * @className: ExcelAnalysisException
 * @author: Kyle.Y.Li
 * @since 1.0.0 2020-04-4/29/2020 16:18
 */
public class ExcelAnalysisException extends RuntimeException {
    public ExcelAnalysisException(){}

    public ExcelAnalysisException(String message) {
        super(message);
    }

    public ExcelAnalysisException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelAnalysisException(Throwable cause) {
        super(cause);
    }
}