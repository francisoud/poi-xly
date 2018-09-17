package com.github.poi.xly.validation;

import java.util.List;
import java.util.function.Predicate;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

/**
 * Validate if value of the cell is present in the {@link #getListOfValues()}.
 * 
 * @see https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/DataValidationHelper.html#createExplicitListConstraint-java.lang.String:A-
 */
public abstract class ExplicitListConstraint implements ExistConstraint<String> {

    private List<String> listOfValues;

    @Override
    public boolean exists(String value) {
        final boolean anyMatch = listOfValues.stream().anyMatch(getPredicate(value));
        return anyMatch;
    }

    @Override
    public String format(String value) {
        return value;
    }

    public List<String> getListOfValues() {
        return listOfValues;
    }

    /**
     * Override if more complex logic is necessary
     */
    public Predicate<? super String> getPredicate(String cellValue) {
        return p -> p.equals(cellValue);
    }

    public void setListOfValues(List<String> listOfValues) {
        this.listOfValues = listOfValues;
    }

    @Override
    public String transform(String cellValue) {
        return cellValue;
    }

    /**
     * Validate if cell value in in the predefine list of values.
     */
    @Override
    public String validate(Cell cell) {
        final DataFormatter dataFormatter = new DataFormatter();
        final String cellValue = dataFormatter.formatCellValue(cell);
        if (cellValue.isEmpty()) {
            return null;
        }
        if (exists(cellValue)) {
            return null;
        }
        return String.format("%s is not a valid value", cellValue);
    }
}