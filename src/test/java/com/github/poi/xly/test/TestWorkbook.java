package com.github.poi.xly.test;

import static com.github.poi.xly.Colors.GREY;
import static com.github.poi.xly.Colors.WHITE;

import java.util.List;

import com.github.poi.xly.annotation.XLYColumn;
import com.github.poi.xly.annotation.XLYSheet;
import com.github.poi.xly.annotation.XLYWorkbook;

@XLYWorkbook
public class TestWorkbook {

    public static final String AGENCIES_HEADER = "Agencies";
    public static final String EMAIL_HEADER = "Email";

    @XLYSheet(name = "Bananas", type = TestBananas.class, rowValidator = TestFlownRowConstraint.class, columns = {
            @XLYColumn(field = "agencies", headerTitle = AGENCIES_HEADER, headerForeground = GREY),
            @XLYColumn(field = "creationDate", mandatory = true, datePattern = "dd/MM/YYYY", headerTitle = "Creation date", headerForeground = GREY, headerFont = WHITE),
            @XLYColumn(field = "quantity", headerTitle = "Quantity", headerForeground = GREY),
            @XLYColumn(field = "revenue", headerTitle = "Revenue", headerForeground = GREY),
            @XLYColumn(field = "email", pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,6}$", headerTitle = EMAIL_HEADER, headerForeground = GREY),
            @XLYColumn(field = "origin", cellValidator = TestOriginConstraint.class, headerTitle = "Origin", headerForeground = GREY),
            @XLYColumn(field = "destination", headerTitle = "Destination", headerForeground = GREY) })
    private List<TestBananas> bananas;

    @XLYSheet(name = "Scenario", type = TestScenario.class, columns = {
            @XLYColumn(field = "name", mandatory = true, headerTitle = "Name", headerForeground = GREY) })
    private List<TestScenario> scenarios;

    public List<TestBananas> getBananas() {
        return bananas;
    }

    public void setBananas(List<TestBananas> bananas) {
        this.bananas = bananas;
    }

    public List<TestScenario> getScenarios() {
        return scenarios;
    }

    public void setScenarios(List<TestScenario> scenarios) {
        this.scenarios = scenarios;
    }
}
