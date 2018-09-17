package com.github.poi.xly.validation;

import java.util.Collections;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * A Constraint that doesn't require any 'constraint'.<br/>
 * Use as default value for rowValidator and cellValidator annotation.
 */
public class NoContraint implements Constraint, RowConstraint {

    @Override
    public Set<String> columnsHeaders() {
        return new HashSet<>();
    }

    @Override
    public String validate(Cell cell) {
        // never return any error message (always return null).
        return null;
    }

    @Override
    public Map<Integer, String> validate(Iterator<Row> rowIterator) {
        // never return any error message (always return emty).
        return Collections.emptyMap();
    }
}
