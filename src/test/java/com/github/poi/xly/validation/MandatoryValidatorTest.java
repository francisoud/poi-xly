package com.github.poi.xly.validation;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.mockito.Mockito.when;

import org.junit.Before;
import org.junit.Test;
import org.mockito.Mockito;

import com.github.poi.xly.annotation.XLYColumn;

public class MandatoryValidatorTest {
    private MandatoryValidator validator;

    @Before
    public void setup() {
        validator = new MandatoryValidator();
    }

    @Test
    public void testValidate_nullCellValue() {
        final String msg = validator.validate(getCellContext(null, true));
        assertNotNull(msg);
        assertEquals("Field required", msg);
    }

    @Test
    public void testValidate_emptyCellValue() {
        final String msg = validator.validate(getCellContext("", true));
        assertNotNull(msg);
        assertEquals("Field required", msg);
    }

    @Test
    public void testValidate_notMandatory() {
        final String msg = validator.validate(getCellContext("some value", false));
        assertNull(msg);
    }

    @Test
    public void testValidate() {
        final String msg = validator.validate(getCellContext("some value", true));
        assertNull(msg);
    }

    private CellContext getCellContext(String value, boolean mandatory) {
        final CellContext cellContext = new CellContext();
        cellContext.setCellValue(value);
        final XLYColumn xlyColumn = Mockito.mock(XLYColumn.class);
        when(xlyColumn.mandatory()).thenReturn(mandatory);
        cellContext.setXlyColumn(xlyColumn);
        return cellContext;
    }
}
