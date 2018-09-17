package com.github.poi.xly.test;

import org.apache.poi.ss.usermodel.Cell;

import com.github.poi.xly.validation.Constraint;

public class TestOriginConstraint implements Constraint {

    @Override
    public String validate(Cell cell) {
        final String cellValue = cell.getStringCellValue();
        if ("FR".equals(cellValue) || "AL".equals(cellValue)) {
            return null;
        }
        return "Origin not exists";
    }
}
