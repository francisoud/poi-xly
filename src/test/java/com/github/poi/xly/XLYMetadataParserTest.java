package com.github.poi.xly;

import static com.github.poi.xly.Colors.GREY;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.junit.Test;

import com.github.poi.xly.annotation.XLYColumn;
import com.github.poi.xly.annotation.XLYSheet;
import com.github.poi.xly.annotation.XLYWorkbook;
import com.github.poi.xly.test.TestBananas;
import com.github.poi.xly.test.TestScenario;
import com.github.poi.xly.test.TestWorkbook;
import com.github.poi.xly.test.XLYFactory;

public class XLYMetadataParserTest {

    private static final int BANANAS_SHEET_INDEX = 0;
    private static final int SCENARIO_SHEET_INDEX = 1;
    private XLYMetadataParser xlyMetadataParser;

    /**
     * Test XLYMetadataParser constructor when passing null workbook.
     */
    @Test(expected = IllegalArgumentException.class)
    public void testXLYMetadataParser_nullWorkbook() {
        xlyMetadataParser = new XLYMetadataParser(null);
    }

    /**
     * Test XLYMetadataParser constructor when passing a workbook object
     * without @XLYWorbook annotation.
     */
    @Test(expected = IllegalArgumentException.class)
    public void testXLYMetadataParser_noWorkbookAnnotation() {
        xlyMetadataParser = new XLYMetadataParser(new Object());
    }

    @Test
    public void testGetSheets() {
        final TestWorkbook workbook = XLYFactory.getWorkbook();
        xlyMetadataParser = new XLYMetadataParser(workbook);
        final List<XLYSheet> sheets = xlyMetadataParser.getSheets();
        assertSheets(sheets);
    }

    @Test
    public void testGetValues() {
        final TestWorkbook xlyWorkbook = XLYFactory.getWorkbook();
        xlyMetadataParser = new XLYMetadataParser(xlyWorkbook);
        final List<XLYSheet> sheets = xlyMetadataParser.getSheets();
        assertSheets(sheets);
        final List<?> bean = xlyMetadataParser.getValues(xlyWorkbook, sheets.get(0));
        assertNotNull(bean);
    }

    @Test(expected = IllegalStateException.class)
    public void testGetValues_invalidWorkbook() {
        final InvalidWorkbookTestXLY xlyWorkbook = new InvalidWorkbookTestXLY();
        xlyMetadataParser = new XLYMetadataParser(xlyWorkbook);
        final List<XLYSheet> sheets = xlyMetadataParser.getSheets();
        xlyMetadataParser.getValues(xlyWorkbook, sheets.get(0));
    }

    @Test
    public void testGetField() {
        final TestWorkbook workbook = XLYFactory.getWorkbook();
        xlyMetadataParser = new XLYMetadataParser(workbook);
        final XLYSheet xlySheet = getFlownSheet();
        final Field field = xlyMetadataParser.getField(xlySheet);
        assertNotNull(field);
        assertEquals("bananas", field.getName());
    }

    @Test
    public void testGetDataFields_TestFlownData() {
        final TestWorkbook workbook = XLYFactory.getWorkbook();
        xlyMetadataParser = new XLYMetadataParser(workbook);
        final Map<String, Field> dataFields = xlyMetadataParser.getDataFields(TestBananas.class,
                getFlownSheet().columns());
        assertNotNull(dataFields);
        final Set<String> fieldsNames = dataFields.keySet();
        assertEquals(7, fieldsNames.size());
        assertNotNull(dataFields.get("agencies"));
        assertNotNull(dataFields.get("creationDate"));
        assertNotNull(dataFields.get("quantity"));
        assertNotNull(dataFields.get("revenue"));
        assertNotNull(dataFields.get("email"));
        assertNotNull(dataFields.get("origin"));
        assertNotNull(dataFields.get("destination"));
    }

    @Test
    public void testGetDataFields_TestScenario() {
        final TestWorkbook workbook = XLYFactory.getWorkbook();
        xlyMetadataParser = new XLYMetadataParser(workbook);
        final Map<String, Field> dataFields = xlyMetadataParser.getDataFields(TestScenario.class,
                getScenarionSheet().columns());
        assertNotNull(dataFields);
        final Set<String> fieldsNames = dataFields.keySet();
        assertEquals(1, fieldsNames.size());
        assertNotNull(dataFields.get("name"));
    }

    private XLYSheet getFlownSheet() {
        final List<XLYSheet> sheets = xlyMetadataParser.getSheets();
        final XLYSheet xlySheet = sheets.get(BANANAS_SHEET_INDEX);
        return xlySheet;
    }

    private XLYSheet getScenarionSheet() {
        final List<XLYSheet> sheets = xlyMetadataParser.getSheets();
        final XLYSheet xlySheet = sheets.get(SCENARIO_SHEET_INDEX);
        return xlySheet;
    }

    private void assertSheets(List<XLYSheet> sheets) {
        assertNotNull(sheets);
        assertEquals(2, sheets.size());
        final XLYSheet flownDataSheet = sheets.get(BANANAS_SHEET_INDEX);
        assertFlownSheet(flownDataSheet);
        final XLYSheet scenarioSheet = sheets.get(SCENARIO_SHEET_INDEX);
        assertScenarioSheet(scenarioSheet);
        for (XLYSheet xlySheet : sheets) {
            // make sure we accidentally didn't put a null in the list
            assertNotNull(xlySheet);
        }
    }

    private void assertFlownSheet(XLYSheet bananasSheet) {
        assertEquals("Bananas", bananasSheet.name());
        final XLYColumn[] bananasHeaders = bananasSheet.columns();
        assertEquals(7, bananasHeaders.length);
        // assert one of the XLYColumn (not testing all of them)
        final XLYColumn creationDate = bananasHeaders[1];
        assertCreationDate(creationDate);
        final XLYColumn origin = bananasHeaders[5];
        assertOrigin(origin);
    }

    private void assertCreationDate(final XLYColumn creationDate) {
        assertEquals("creationDate", creationDate.field());
        assertEquals("Creation date", creationDate.headerTitle());
        assertEquals(Colors.GREY, creationDate.headerForeground());
        assertEquals(Colors.WHITE, creationDate.headerFont());
    }

    private void assertOrigin(XLYColumn origin) {
        assertEquals("origin", origin.field());
        assertEquals("Origin", origin.headerTitle());
        assertEquals(Colors.GREY, origin.headerForeground());
    }

    private void assertScenarioSheet(XLYSheet scenarioSheet) {
        assertEquals("Scenario", scenarioSheet.name());
        final XLYColumn[] scenarioHeaders = scenarioSheet.columns();
        assertEquals(1, scenarioHeaders.length);
    }

    @XLYWorkbook
    public class InvalidWorkbookTestXLY {
        /**
         * Instead of a List<TestScenario> we have a simple POJO: TestScenario
         */
        @XLYSheet(name = "Scenario", type = TestScenario.class, columns = {
                @XLYColumn(field = "name", headerTitle = "Name", headerForeground = GREY) })
        private TestScenario scenarii = new TestScenario();

        public TestScenario getScenarii() {
            return scenarii;
        }

        public void setScenarii(TestScenario scenarii) {
            this.scenarii = scenarii;
        }
    }
}
