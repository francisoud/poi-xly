package com.github.poi.xly.validation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

/**
 * A method to check that a value in excel exist 'somewhere'. <br/>
 * Implementation can use a predefine list of values or most usually check
 * existence in the database using Repository.
 */
public interface ExistConstraint<T> extends Constraint {

    boolean exists(T value);

    /**
     * Format the entity to be included in error message.
     */
    String format(T value);

    T transform(String cellValue);

    @Override
    default String validate(Cell cell) {
        final DataFormatter dataFormatter = new DataFormatter();
        final String cellValue = dataFormatter.formatCellValue(cell);
        if (cellValue.isEmpty()) {
            return null;
        }
        final T entity = transform(cellValue);
        if (exists(entity)) {
            return null;
        }
        return String.format("%s is not a valid value", format(entity));
    }
}
