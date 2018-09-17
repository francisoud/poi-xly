package com.github.poi.xly.validation;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DefaultConstraintLocator implements ConstraintLocator {

    private static final Logger logger = LoggerFactory.getLogger(DefaultConstraintLocator.class);

    @Override
    public Constraint getConstraint(Class<? extends Constraint> validatorClass) {
        try {
            return validatorClass.newInstance();
        } catch (ReflectiveOperationException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e.getMessage(), e);
        }
    }

    @Override
    public RowConstraint getRowConstraint(Class<? extends RowConstraint> validatorClass) {
        try {
            return validatorClass.newInstance();
        } catch (ReflectiveOperationException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e.getMessage(), e);
        }
    }
}
