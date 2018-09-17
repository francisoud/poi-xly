package com.github.poi.xly;

@SuppressWarnings("serial")
public class XLYException extends RuntimeException {

    public enum XLYError {
        MISSING_SHEET,
        /**
         * Importing an excel in protected mode make poi crash with "Zip bomb
         * detected" unrelated message :(
         * https://support.office.com/en-us/article/What-is-Protected-View-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
         */
        PROTECTED_VIEW_ENABLE, UNEXCEPTED_ERROR
    }

    private XLYError xlyError = XLYError.UNEXCEPTED_ERROR;

    public XLYException(Exception root) {
        super(root.getMessage(), root);
    }

    /**
     * @param message
     */
    public XLYException(XLYError message) {
        super(message.name());
        this.xlyError = message;
    }

    /**
     * @param message
     * @param cause
     */
    public XLYException(XLYError message, Throwable cause) {
        super(message.name(), cause);
        this.xlyError = message;
    }

    public XLYError getXlyError() {
        return xlyError;
    }
}
