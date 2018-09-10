package com.unaware.poi.excel.exception;

/**
 * @author Unaware
 * @Description: the exception for wrong parameter
 * @Title: ParameterException
 * @ProjectName excel
 * @date 2018/7/25 10:19
 */
public class ParameterException extends RuntimeException {
    public ParameterException() {
        super();
    }

    public ParameterException(String msg) {
        super(msg);
    }

    public ParameterException(Exception e) {
        super(e);
    }

    public ParameterException(String msg, Exception e) {
        super(msg, e);
    }
}
