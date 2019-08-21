package com.lwb.easy.excel.exception;

/**
 * @author liuweibo
 * @date 2019/8/20
 */
public class ExcelException extends RuntimeException {

    public ExcelException() {
        super();
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelException(Throwable cause) {
        super(cause);
    }

}
