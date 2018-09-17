package com.github.poi.xly.test;

import org.apache.poi.ss.usermodel.Cell;

import com.github.poi.xly.validation.Constraint;

public class MyCustomConstraint implements Constraint {
    @Override
    public String validate(Cell cell) {
        return "my custom message";
    }
}
