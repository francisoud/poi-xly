package com.github.poi.xly.validation;

import org.junit.Test;

import com.github.poi.xly.test.WorkbookTest;

public class ExistConstraintTest extends WorkbookTest {

    @Test
    public void testValidate() {
        final ExistConstraint<String> existConstraint = new MyExistConstraint();
        existConstraint.validate(getCell("a value"));
    }

    @Test
    public void testValidate_KO() {
        final ExistConstraint<String> existConstraint = new MyExistConstraint();
        existConstraint.validate(getCell("NOT AN EXISTING value"));
    }

    private class MyExistConstraint implements ExistConstraint<String> {

        @Override
        public boolean exists(String value) {
            return "a value".equals(value);
        }

        @Override
        public String transform(String cellValue) {
            return cellValue;
        }

        @Override
        public String format(String value) {
            return value;
        }
    }
}
