package com.github.poi.xly.validation;

import java.util.ArrayList;
import java.util.List;

import com.github.poi.xly.annotation.XLYColumn;

public class UserDefineValidator implements CellValidator {

    private final ConstraintLocator constraintLocator;

    private List<String> possibleValues = new ArrayList<>();

    public UserDefineValidator(ConstraintLocator constraintLocator) {
        this.constraintLocator = constraintLocator;
    }

    public List<String> getPossibleValues() {
        return possibleValues;
    }

    public void setPossibleValues(List<String> possibleValues) {
        this.possibleValues = possibleValues;
    }

    @Override
    public String validate(CellContext cellContext) {
        final XLYColumn xlyColumn = cellContext.getXlyColumn();
        final Class<? extends Constraint> validatorClass = xlyColumn.cellValidator();
        if (NoContraint.class.equals(validatorClass)) {
            return null;
        }
        final Constraint validator = constraintLocator.getConstraint(validatorClass);
        if (validator instanceof ExplicitListConstraint) {
            final ExplicitListConstraint explicitListConstraint = (ExplicitListConstraint) validator;
            explicitListConstraint.setListOfValues(possibleValues);
        }
        return validator.validate(cellContext.getCell());
    }
}
