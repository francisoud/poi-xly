package com.github.poi.xly.validation;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import java.lang.reflect.Field;
import java.time.DayOfWeek;

import org.junit.Test;

public class EnumValidatorTest {

    /** An enum use for thsi test */
    public DayOfWeek dayOfWeek;

    @Test
    public void testValidate() throws Exception {
        final EnumValidator validator = new EnumValidator();
        final CellContext cellContext = getCellContext(DayOfWeek.FRIDAY.name());
        assertNull(validator.validate(cellContext));
    }

    @Test
    public void testValidate_KO() throws Exception {
        final EnumValidator validator = new EnumValidator();
        final CellContext cellContext = getCellContext("NO_A_VALID_WEEK_DAY");
        final String msg = validator.validate(cellContext);
        assertNotNull(msg);
        assertEquals(
                "NO_A_VALID_WEEK_DAY not exists. Did you mean : MONDAY or TUESDAY or WEDNESDAY or THURSDAY or FRIDAY or SATURDAY or SUNDAY ?",
                msg);
    }

    @Test
    public void testValidate_null() throws Exception {
        final EnumValidator validator = new EnumValidator();
        final CellContext cellContext = getCellContext(null);
        assertNull(validator.validate(cellContext));
    }

    private CellContext getCellContext(String dayOfWeek) throws NoSuchFieldException {
        final CellContext cellContext = new CellContext();
        cellContext.setCellValue(dayOfWeek);
        Field field = this.getClass().getField("dayOfWeek");
        cellContext.setField(field);
        return cellContext;
    }
}
