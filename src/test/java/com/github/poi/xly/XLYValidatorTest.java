package com.github.poi.xly;

import static com.github.poi.xly.test.XLYAssert.assertCellsWithErrorStyle;
import static com.github.poi.xly.test.XLYAssert.assertFlownSheet;
import static com.github.poi.xly.test.XLYAssert.assertScenarioSheet;
import static com.github.poi.xly.test.XLYAssert.toWorkbook;
import static com.github.poi.xly.test.XLYFactory.BANANAS_DATA_SHEET_NAME;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.mockito.Matchers.any;
import static org.mockito.Matchers.anyInt;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.github.poi.xly.annotation.XLYSheet;
import com.github.poi.xly.annotation.XLYWorkbook;
import com.github.poi.xly.test.TestBananas;
import com.github.poi.xly.test.TestWorkbook;
import com.github.poi.xly.test.XLYFactory;
import com.github.poi.xly.validation.ConstraintLocator;
import com.github.poi.xly.validation.DefaultConstraintLocator;

public class XLYValidatorTest {

    private ConstraintLocator constraintLocator = new DefaultConstraintLocator();

    private static final String[] EXPECTED_BANANAS_ROWS = {
            "FR;6/8/18;2;INVALID_REVENUE;INVALID_EMAIL;INVALID_ORIGIN;FR;Invalid cell value. Value doesn't match with this pattern : ^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,6}$Origin not exists;",
            "FR;08/32/2018;INVALID_OPPORTUNITY;2;email2@corp2.com;AL;INVALID_DESTINATION;",
            "FR;;2;2;date.mandatory.but.empty@corp1.com;FR;AL;Field required;",
            "FR;6/8/18;2;2;valid.line@corp1.com;FR;AL;" };

    private static final String[] EXPECTED_BANANAS_ROWS_DUP = { "FR;6/8/18;2;2;duplicated.line@corp1.com;FR;AL;",
            "FR;6/8/18;2;2;duplicated.line@corp1.com;FR;AL;Duplicate line;" };

    private static final String[] EXPECTED_SCENARIO_ROWS = { "SCENARIO;" };

    private XLYValidator xlyValidator;
    private ByteArrayOutputStream outputStream;

    @Before
    public void setup() {
        XLYFactory.setup();
        // replace with a FileOutputStream if you want to open the result file
        // in excel
        // outputStream = new FileOutputStream("/tmp/XLYValidationTest.xlsx");
        outputStream = new ByteArrayOutputStream();
    }

    @After
    public void tearDown() throws IOException {
        outputStream.close();
    }

    /**
     * Test IllegalArgumentException if null workbook object.
     */
    @Test(expected = IllegalArgumentException.class)
    public void testXLYValidator_nullConstraintLocator() {
        xlyValidator = new XLYValidator(null);
    }

    @Test
    public void testIsValid() {
        xlyValidator = new XLYValidator(constraintLocator);
        xlyValidator.setWorkbookClass(TestWorkbook.class);
        final InputStream inputStream = XLYFactory.getBananasOK();
        final boolean isValid = xlyValidator.isValid(inputStream, outputStream);
        assertTrue(isValid);
    }

    /**
     * Test IllegalArgumentException if null workbookClass object.
     */
    @Test(expected = IllegalArgumentException.class)
    public void testValidate_nullWorkbookClass() {
        xlyValidator = new XLYValidator(constraintLocator);
        final InputStream inputStream = XLYFactory.getBananasOK();
        xlyValidator.validate(inputStream, outputStream);
    }

    @Test
    public void testValidate() throws IOException {
        xlyValidator = new XLYValidator(constraintLocator);
        xlyValidator.setWorkbookClass(TestWorkbook.class);
        final InputStream inputStream = XLYFactory.getBananasOK();
        final OutputStream mockOutputStream = mock(OutputStream.class);
        final Set<String> violations = xlyValidator.validate(inputStream, mockOutputStream);
        assertNotNull(violations);
        final String msg = "validation result was suppose to be empty but got: " + String.join(",", violations);
        assertTrue(msg, violations.isEmpty());
        // make sure if validation is ok => don't write anything to output
        verify(mockOutputStream, never()).write(any());
        verify(mockOutputStream, never()).write(anyInt());
        verify(mockOutputStream, never()).write(any(), anyInt(), anyInt());
    }

    @Test
    public void testValidate_KO_cellErrors() throws IOException {
        xlyValidator = new XLYValidator(constraintLocator);
        xlyValidator.setWorkbookClass(TestWorkbook.class);
        final InputStream inputStream = XLYFactory.getBananasKO_cellErrors();
        final Set<String> violations = xlyValidator.validate(inputStream, outputStream);
        assertNotNull(violations);
        assertFalse(violations.isEmpty());
        assertViolations(violations);
        final XSSFWorkbook generatedWorkbook = toWorkbook(outputStream);
        assertEquals(2, generatedWorkbook.getNumberOfSheets());
        assertFlownSheet(generatedWorkbook, EXPECTED_BANANAS_ROWS);
        assertScenarioSheet(generatedWorkbook, EXPECTED_SCENARIO_ROWS);
        assertCellsWithErrorStyle(generatedWorkbook, 0, "E2", "F2", "H2", "B4", "H4");
        final XSSFColor tabColor = generatedWorkbook.getSheet(BANANAS_DATA_SHEET_NAME).getTabColor();
        assertEquals(XLYFormatter.RED, tabColor);
        generatedWorkbook.close();
    }

    @Test
    public void testValidate_KO_duplicatedLines() throws IOException {
        xlyValidator = new XLYValidator(constraintLocator);
        xlyValidator.setWorkbookClass(TestWorkbook.class);
        final InputStream inputStream = XLYFactory.getBananasKO_duplicatedLines();
        final Set<String> violations = xlyValidator.validate(inputStream, outputStream);
        assertNotNull(violations);
        assertFalse(violations.isEmpty());
        final String msg = String.join(",", violations);
        assertTrue(msg, violations.contains("line:3 - error: Duplicate line"));
        assertEquals(msg, 1, violations.size());
        final XSSFWorkbook generatedWorkbook = toWorkbook(outputStream);
        assertEquals(2, generatedWorkbook.getNumberOfSheets());
        assertFlownSheet(generatedWorkbook, EXPECTED_BANANAS_ROWS_DUP);
        assertScenarioSheet(generatedWorkbook, EXPECTED_SCENARIO_ROWS);
        assertCellsWithErrorStyle(generatedWorkbook, 0, "H3");
        final XSSFColor tabColor = generatedWorkbook.getSheet(BANANAS_DATA_SHEET_NAME).getTabColor();
        assertEquals(XLYFormatter.RED, tabColor);
        generatedWorkbook.close();
    }

    @Test(expected = XLYException.class)
    public void testValidate_KO_missingSheet() {
        xlyValidator = new XLYValidator(constraintLocator);
        xlyValidator.setWorkbookClass(TestWorkbook.class);
        final InputStream inputStream = XLYFactory.getBananasKO_missingSheet();
        xlyValidator.validate(inputStream, outputStream);
    }

    /**
     * Another way to test for a missing sheet than @link
     * {@link XLYValidatorTest#testValidate_KO_missingSheet()}
     */
    @Test(expected = XLYException.class)
    public void testValidate_KO_missingSheet2() {
        xlyValidator = new XLYValidator(constraintLocator);
        xlyValidator.setWorkbookClass(MissingSheetWorkbook.class);
        final InputStream inputStream = XLYFactory.getBananasKO_missingSheet();
        xlyValidator.validate(inputStream, outputStream);
    }

    @Test
    public void testGetWorkbookClass() {
        xlyValidator = new XLYValidator(constraintLocator);
        xlyValidator.setWorkbookClass(TestWorkbook.class);
        assertEquals(TestWorkbook.class, xlyValidator.getWorkbookClass());
    }

    private void assertViolations(Set<String> violations) {
        final String msg = String.join(",", violations);
        assertTrue(msg, violations.contains("Field required"));
        assertTrue(msg, violations.contains(
                "Invalid cell value. Value doesn't match with this pattern : ^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,6}$"));
        assertTrue(msg, violations.contains("Origin not exists"));
        assertEquals(msg, 3, violations.size());
    }

    @XLYWorkbook
    private class MissingSheetWorkbook extends TestWorkbook {
        @XLYSheet(name = "MISSING SHEET NAME", toImport = true, type = TestBananas.class, columns = {})
        private List<TestBananas> missing;

        @SuppressWarnings("unused")
        public List<TestBananas> getMissing() {
            return missing;
        }

        @SuppressWarnings("unused")
        public void setMissing(List<TestBananas> missing) {
            this.missing = missing;
        }
    }
}
