package com.github.poi.xly.test;

import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

//import org.hibernate.ScrollableResults;

public final class XLYFactory {

    public static final String SCENARIO_SHEET_NAME = "Scenario";
    public static final String BANANAS_DATA_SHEET_NAME = "Bananas";
    public static final String BANANAS_DATA_OK = "bananasOK.xlsx";
    public static final String BANANAS_DATA_CELLS_KO = "bananasKO_cellErrors.xlsx";

    private static SimpleDateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd");

    /**
     * No constructor needed. Only static methods.
     */
    private XLYFactory() {
    }

    public static TestWorkbook getWorkbook() {
        final TestWorkbook workbook = new TestWorkbook();
        workbook.setBananas(getBananas());
        workbook.setScenarios(getScenario());
        return workbook;
    }

    public static TestBananas getBananas(String origin, String destination, String email) {
        final TestBananas bananas = new TestBananas();
        bananas.setCreationDate(getDate());
        bananas.setQuantity(2);
        bananas.setRevenue(2d);
        bananas.setAgencies("FR");
        bananas.setEmail(email);
        bananas.setOrigin(origin);
        bananas.setDestination(destination);
        return bananas;
    }

    /**
     * set headless mode to avoid poi font errors during unit tests.<br/>
     * 'java.lang.NoClassDefFoundError: Could not initialize class
     * sun.awt.X11GraphicsEnvironment'
     */
    public static void setup() {
        System.setProperty("java.awt.headless", "true");
    }

    public static InputStream getBananasOK() {
        return getFile("/xly/" + BANANAS_DATA_OK);
    }

    public static InputStream getBananasKO_cellErrors() {
        return getFile("/xly/" + BANANAS_DATA_CELLS_KO);
    }

    public static InputStream getBananasKO_duplicatedLines() {
        return getFile("/xly/bananasKO_duplicatedLines.xlsx");
    }

    public static InputStream getBananasKO_missingSheet() {
        return getFile("/xly/bananasKO_missingSheet.xlsx");
    }

    public static List<TestScenario> getScenario() {
        final List<TestScenario> scenarios = new ArrayList<>();
        final TestScenario scenario = getScenarii();
        scenarios.add(scenario);
        return scenarios;
    }

    public static List<TestBananas> getBananas() {
        final List<TestBananas> bananas = new ArrayList<>();
        final TestBananas firstTrip = getBananas("FR", "AL", "email1@corp1.com");
        bananas.add(firstTrip);
        final TestBananas returnTrip = getBananas("AL", "FR", "email2@corp2.com");
        bananas.add(returnTrip);
        return bananas;
    }

    private static TestScenario getScenarii() {
        final TestScenario scenario = new TestScenario();
        scenario.setName("SCENARIO");
        return scenario;
    }

    private static Date getDate() {
        try {
            return dateFormatter.parse("2018-06-08");
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }

    private static InputStream getFile(String filename) {
        final InputStream in = XLYFactory.class.getResourceAsStream(filename);
        assert (in != null);
        return in;
    }
}
