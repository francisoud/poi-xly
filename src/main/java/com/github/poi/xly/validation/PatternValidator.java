package com.github.poi.xly.validation;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PatternValidator implements CellValidator {

    @Override
    public String validate(CellContext cellContext) {
        final String cellValue = cellContext.getCellValue();
        if (cellValue == null) {
            return null;
        }
        final String pattern = cellContext.getXlyColumn().pattern();
        final Pattern validatorPattern = Pattern.compile(pattern, Pattern.CASE_INSENSITIVE);
        final Matcher matcher = validatorPattern.matcher(cellValue);
        if (!matcher.find()) {
            return String.format("Invalid cell value. Value doesn't match with this pattern : %s", pattern);
        }
        return null;
    }
}
