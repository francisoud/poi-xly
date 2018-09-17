package com.github.poi.xly;

import static com.github.poi.xly.test.XLYAssert.assertFlownSheet;
import static com.github.poi.xly.test.XLYAssert.assertScenarioSheet;
import static com.github.poi.xly.test.XLYAssert.toWorkbook;
import static org.junit.Assert.assertEquals;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.github.poi.xly.test.TestWorkbook;
import com.github.poi.xly.test.XLYFactory;

public class XLYExporterTest {

    public XLYExporter xlyExporter;

    private ByteArrayOutputStream outputStream;

    private static final String[] EXPECTED_FLOWN_ROWS = { "FR;08/06/2018;2;2;email1@corp1.com;FR;AL;",
            "FR;08/06/2018;2;2;email2@corp2.com;AL;FR;" };

    private static final String[] EXPECTED_SCENARIO_ROWS = { "SCENARIO;" };

    @Before
    public void setup() {
        XLYFactory.setup();
        // replace with a FileOutputStream if you want to open the result file
        // in excel
        // outputStream = new FileOutputStream("/tmp/XLYExporterTest.xlsx");
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
    public void testGenerate_nullWorkbook() {
        xlyExporter = new XLYExporter(null);
        xlyExporter.export(outputStream);
    }

    @Test
    public void testExport() throws Exception {
        final TestWorkbook workbook = XLYFactory.getWorkbook();
        xlyExporter = new XLYExporter(workbook);
        xlyExporter.export(outputStream);
        final XSSFWorkbook generatedWorkbook = toWorkbook(outputStream);
        assertEquals(2, generatedWorkbook.getNumberOfSheets());
        assertFlownSheet(generatedWorkbook, EXPECTED_FLOWN_ROWS);
        assertScenarioSheet(generatedWorkbook, EXPECTED_SCENARIO_ROWS);
        generatedWorkbook.close();
    }
}
