package com.github.poi.xly;

import static org.junit.Assert.assertEquals;

import java.io.IOException;

import org.junit.Test;

import com.github.poi.xly.XLYException.XLYError;

public class XLYExceptionTest {

    @Test
    public void testXLYExceptionXLYErrorThrowable() {
        final String cause = "the cause excpetion message";
        XLYException exception = new XLYException(XLYError.UNEXCEPTED_ERROR, new RuntimeException(cause));
        assertEquals("UNEXCEPTED_ERROR", exception.getMessage());
        assertEquals(XLYError.UNEXCEPTED_ERROR, exception.getXlyError());
        assertEquals(cause, exception.getCause().getMessage());
    }

    @Test
    public void testXLYExceptionXLYError() {
        XLYException exception = new XLYException(XLYError.UNEXCEPTED_ERROR);
        assertEquals("UNEXCEPTED_ERROR", exception.getMessage());
        assertEquals(XLYError.UNEXCEPTED_ERROR, exception.getXlyError());
    }

    @Test
    public void testXLYException() {
        final IOException anException = new IOException("an exception message");
        XLYException exception = new XLYException(anException);
        assertEquals("an exception message", exception.getMessage());
        assertEquals(XLYError.UNEXCEPTED_ERROR, exception.getXlyError());
    }
}
