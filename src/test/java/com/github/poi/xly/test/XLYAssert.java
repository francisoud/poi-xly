package com.github.poi.xly.test;

import static com.github.poi.xly.test.XLYFactory.BANANAS_DATA_SHEET_NAME;
import static com.github.poi.xly.test.XLYFactory.SCENARIO_SHEET_NAME;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.github.poi.xly.Colors;
import com.github.poi.xly.XLYFormatter;
import com.github.poi.xly.XLYFormatterTest;

/**
 * Test utility method to assert the content of excel files. <br/>
 * Highly dependent of content produce by {@link XLYFactory}
 */
public class XLYAssert {

    public static XSSFWorkbook toWorkbook(ByteArrayOutputStream outputStream) throws IOException {
        final byte[] byteArray = outputStream.toByteArray();
        assertNotNull(byteArray);
        final XSSFWorkbook generatedWorkbook = new XSSFWorkbook(new ByteArrayInputStream(byteArray));
        return generatedWorkbook;
    }

    public static void assertFlownSheet(final XSSFWorkbook generatedWorkbook, String[] expectedRows) {
        final XSSFSheet flownSheet = generatedWorkbook.getSheet(BANANAS_DATA_SHEET_NAME);
        assertNotNull(flownSheet);
        assertEquals(expectedRows.length, flownSheet.getLastRowNum());
        final XSSFRow flownHeader = flownSheet.getRow(0);
        assertFlownHeader(generatedWorkbook, flownHeader);
        for (int i = 0; i < expectedRows.length; i++) {
            final int rownum = i + 1; // skip header
            assertRowContent(expectedRows[i], flownSheet.getRow(rownum));
        }
    }

    public static void assertScenarioSheet(XSSFWorkbook generatedWorkbook, String[] expectedRows) {
        final XSSFSheet scenarioSheet = generatedWorkbook.getSheet(SCENARIO_SHEET_NAME);
        assertNotNull(scenarioSheet);
        assertEquals(1, scenarioSheet.getLastRowNum());
        final XSSFRow scenarioHeader = scenarioSheet.getRow(0);
        assertEquals("Name", scenarioHeader.getCell(0).getStringCellValue());
        for (int i = 0; i < expectedRows.length; i++) {
            final int rownum = i + 1; // skip header
            assertRowContent(expectedRows[i], scenarioSheet.getRow(rownum));
        }
    }

    /**
     * @param references
     *            for example B12, C3 etc.
     */
    public static void assertCellsWithErrorStyle(XSSFWorkbook generatedWorkbook, int sheetIndex, String... references) {
        for (String reference : references) {
            final CellReference ref = new CellReference(reference);
            final Cell cell = generatedWorkbook.getSheetAt(sheetIndex).getRow(ref.getRow()).getCell(ref.getCol());
            assertEquals(XLYFormatter.RED_INDEX, cell.getCellStyle().getFillForegroundColor());
            assertEquals(XLYFormatter.RED_INDEX, cell.getCellStyle().getFillBackgroundColor());
        }
    }

    private static void assertFlownHeader(final XSSFWorkbook workbook, XSSFRow flownHeader) {
        assertEquals("Agencies", flownHeader.getCell(0).getStringCellValue());
        final XSSFCell creationDateHeader = flownHeader.getCell(1);
        assertEquals("Creation date", creationDateHeader.getStringCellValue());
        XLYFormatterTest.assertColors(Colors.GREY, Colors.WHITE, workbook, creationDateHeader.getCellStyle());
        assertEquals("Quantity", flownHeader.getCell(2).getStringCellValue());
        assertEquals("Revenue", flownHeader.getCell(3).getStringCellValue());
        assertEquals("Email", flownHeader.getCell(4).getStringCellValue());
        assertEquals("Origin", flownHeader.getCell(5).getStringCellValue());
        assertEquals("Destination", flownHeader.getCell(6).getStringCellValue());
    }

    /**
     * Recreate the excel row as a string to be able to assert the content
     * easily
     */
    private static void assertRowContent(String expected, XSSFRow row) {
        DataFormatter dataFormatter = new DataFormatter(Locale.FRANCE);
        final StringBuilder sb = new StringBuilder();
        final Iterator<Cell> iterator = row.cellIterator();
        while (iterator.hasNext()) {
            final Cell cell = iterator.next();
            sb.append(dataFormatter.formatCellValue(cell));
            sb.append(";");
        }
        assertEquals(expected, sb.toString().replaceAll("[\n\r]", ""));
    }
}
