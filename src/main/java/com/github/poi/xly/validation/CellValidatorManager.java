package com.github.poi.xly.validation;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

import com.github.poi.xly.XLYFormatter;
import com.github.poi.xly.annotation.XLYColumn;

/**
 * Responsible of running all CellValidators in the correct order. <br/>
 * 
 * @see https://en.wikipedia.org/wiki/Chain-of-responsibility_pattern
 * @see https://docs.oracle.com/javaee/6/api/javax/servlet/FilterChain.html
 */
public class CellValidatorManager {

    private final List<CellValidator> cellValidators = new ArrayList<>();

    private final XLYFormatter xlyFormatter;

    public CellValidatorManager(ConstraintLocator constraintLocator, XLYFormatter xlyFormatter) {
        this.xlyFormatter = xlyFormatter;
        // configure validators and validation order
        cellValidators.add(new MandatoryValidator());
        cellValidators.add(new UserDefineValidator(constraintLocator));
        cellValidators.add(new PatternValidator());
        cellValidators.add(new EnumValidator());
    }

    /**
     * Get the value of the cell as a formatted string.
     */
    private String getValue(Cell cell) {
        final DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    /**
     * Run all CellValidator on the specific cell.
     */
    public Set<String> validate(Cell cell, XLYColumn xlyColumn, Field field) {
        final Set<String> violations = new HashSet<>();
        final String value = getValue(cell);
        for (CellValidator cellValidator : cellValidators) {
            final CellContext cellContext = new CellContext(cell, value, xlyColumn, field);
            final String errorMessage = cellValidator.validate(cellContext);
            if (errorMessage != null) {
                violations.add(errorMessage);
                xlyFormatter.addErrorMessage(cell, errorMessage);
            }
        }
        return violations;
    }
}
