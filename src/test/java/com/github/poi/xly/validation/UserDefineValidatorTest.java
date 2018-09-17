package com.github.poi.xly.validation;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.mockito.Mockito.when;

import java.util.Arrays;

import org.junit.Before;
import org.junit.Test;
import org.mockito.Mockito;

import com.github.poi.xly.annotation.XLYColumn;
import com.github.poi.xly.test.MyCustomConstraint;
import com.github.poi.xly.test.MyCustomListConstraint;
import com.github.poi.xly.test.WorkbookTest;

public class UserDefineValidatorTest extends WorkbookTest {

    private UserDefineValidator userDefineValidator;

    @Before
    public void createValidator() {
        userDefineValidator = new UserDefineValidator(new DefaultConstraintLocator());
    }

    @Test
    public void testValidate_customConstraint() {
        final CellContext cellContext = initCustomConstraint("xxx");
        final String msg = userDefineValidator.validate(cellContext);
        assertNotNull(msg);
        assertEquals("my custom message", msg);
    }

    @Test
    public void testValidate_listConstraint() {
        final CellContext cellContext = initConstraint("xxx");
        userDefineValidator.setPossibleValues(Arrays.asList("A", "B"));
        final String msg = userDefineValidator.validate(cellContext);
        assertNotNull(msg);
        assertEquals("xxx is not a valid value", msg);
    }

    @Test
    public void testValidate_listConstraint_validValue() {
        final CellContext cellContext = initConstraint("A");
        userDefineValidator.setPossibleValues(Arrays.asList("A", "B"));
        final String msg = userDefineValidator.validate(cellContext);
        assertNull(msg);
    }

    private CellContext initCustomConstraint(String cellValue) {
        final CellContext cellContext = new CellContext();
        cellContext.setCell(getCell(cellValue));
        setConstraint(cellContext, MyCustomConstraint.class);
        return cellContext;
    }

    private CellContext initConstraint(String cellValue) {
        final CellContext cellContext = new CellContext();
        cellContext.setCell(getCell(cellValue));
        setConstraint(cellContext, MyCustomListConstraint.class);
        return cellContext;
    }

    @SuppressWarnings("unchecked")
    private void setConstraint(CellContext cellContext, @SuppressWarnings("rawtypes") Class constraintClass) {
        final XLYColumn xlyColumn = Mockito.mock(XLYColumn.class);
        when(xlyColumn.cellValidator()).thenReturn(constraintClass);
        cellContext.setXlyColumn(xlyColumn);
    }
}
