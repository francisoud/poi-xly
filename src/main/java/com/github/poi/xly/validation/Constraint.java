package com.github.poi.xly.validation;

import org.apache.poi.ss.usermodel.Cell;

/**
 * Interface for cell validation.
 */
public interface Constraint {
    String validate(Cell cell);
}
