package com.github.poi.xly.validation;

import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.github.poi.xly.XLYFormatter;
import com.github.poi.xly.annotation.XLYSheet;

public class RowValidatorManager {

    private final ConstraintLocator constraintLocator;

    private final XLYFormatter xlyFormatter;

    public RowValidatorManager(ConstraintLocator constraintLocator, XLYFormatter xlyFormatter) {
        this.constraintLocator = constraintLocator;
        this.xlyFormatter = xlyFormatter;
    }

    private void addErrorCell(XSSFSheet sheet, Map<Integer, String> violations, Integer rownum) {
        final Row row = sheet.getRow(rownum);
        xlyFormatter.addErrorMessage(row, violations.get(rownum));
    }

    private Set<String> handleViolations(XSSFSheet sheet, Map<Integer, String> rows) {
        final Set<String> violations = new HashSet<>();
        for (Integer rownum : rows.keySet()) {
            // +1 because getRowNum() is 0 based (getRowNum():0 means row:1 in
            // excel)
            final String msg = String.format("line:%s - error: %s", rownum + 1, rows.get(rownum));
            violations.add(msg);
            addErrorCell(sheet, rows, rownum);
        }
        return violations;
    }

    public Set<String> validateRows(XSSFSheet sheet, XLYSheet xlySheet) {
        final Set<String> violations = new HashSet<>();
        final Class<? extends RowConstraint>[] validators = xlySheet.rowValidator();
        for (Class<? extends RowConstraint> validatorClass : validators) {
            final RowConstraint validator = constraintLocator.getRowConstraint(validatorClass);
            final Iterator<Row> rowIterator = sheet.iterator();
            final Map<Integer, String> rowViolations = validator.validate(rowIterator);
            violations.addAll(handleViolations(sheet, rowViolations));
        }
        return violations;
    }
}
