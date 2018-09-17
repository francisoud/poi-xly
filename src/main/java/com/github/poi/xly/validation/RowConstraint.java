package com.github.poi.xly.validation;

import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;

/**
 * Constraint on rows of a specific sheet.
 */
public interface RowConstraint {
    /**
     * A list of headers that will be use to perform row duplication check.
     * <br/>
     * A row will be marked as duplicate if the content of all column listed in
     * those headers are identical.
     */
    public abstract Set<String> columnsHeaders();

    /**
     * @return Map&lt;rowNum, errorMesage&gt; empty if no error.
     */
    Map<Integer, String> validate(Iterator<Row> rowIterator);
}
