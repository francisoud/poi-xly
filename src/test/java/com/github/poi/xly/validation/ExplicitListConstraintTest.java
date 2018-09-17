package com.github.poi.xly.validation;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.util.Arrays;
import java.util.List;
import java.util.function.Predicate;

import org.junit.Test;

import com.github.poi.xly.test.WorkbookTest;

public class ExplicitListConstraintTest extends WorkbookTest {

    private ExplicitListConstraint explicitListConstraint;

    @Test
    public void testGetPredicate() {
        explicitListConstraint = new CustomPredicateExplicitListConstraint();
        explicitListConstraint.setListOfValues(getListOfValues());
        assertNull(explicitListConstraint.validate(getCell("a")));
    }

    @Test
    public void testValidate_nullValue() {
        explicitListConstraint = new SimpleExplicitListConstraint();
        explicitListConstraint.setListOfValues(getListOfValues());
        assertNull(explicitListConstraint.validate(getCell(null)));
    }

    @Test
    public void testValidate() {
        explicitListConstraint = new SimpleExplicitListConstraint();
        explicitListConstraint.setListOfValues(getListOfValues());
        final String msg = explicitListConstraint.validate(getCell("NOT_A_NOR_B_NOR_C"));
        assertEquals("NOT_A_NOR_B_NOR_C is not a valid value", msg);
    }

    private List<String> getListOfValues() {
        return Arrays.asList("A", "B", "C");
    }

    private class SimpleExplicitListConstraint extends ExplicitListConstraint {
    }

    /**
     * compare string using lower case.
     */
    private class CustomPredicateExplicitListConstraint extends ExplicitListConstraint {
        @Override
        public Predicate<? super String> getPredicate(String cellValue) {
            return p -> p.toLowerCase().equals(cellValue);
        }
    }
}
