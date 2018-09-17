package com.github.poi.xly.validation;

public interface ConstraintLocator {

    Constraint getConstraint(Class<? extends Constraint> validatorClass);

    RowConstraint getRowConstraint(Class<? extends RowConstraint> validatorClass);
}
