package com.github.poi.xly.test;

import static com.github.poi.xly.test.TestWorkbook.AGENCIES_HEADER;
import static com.github.poi.xly.test.TestWorkbook.EMAIL_HEADER;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

import com.github.poi.xly.validation.DuplicatedRowConstraint;

public class TestFlownRowConstraint extends DuplicatedRowConstraint {

    @Override
    public Set<String> columnsHeaders() {
        return new HashSet<>(Arrays.asList(AGENCIES_HEADER, EMAIL_HEADER));
    }
}
