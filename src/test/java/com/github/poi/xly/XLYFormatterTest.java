package com.github.poi.xly;

import static com.github.poi.xly.test.XLYAssert.assertCellsWithErrorStyle;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.mockito.Matchers.anyInt;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.times;
import static org.mockito.Mockito.verify;

import java.lang.reflect.Field;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.Before;
import org.junit.Test;

import com.github.poi.xly.annotation.XLYColumn;
import com.github.poi.xly.annotation.XLYSheet;
import com.github.poi.xly.annotation.XLYWorkbook;
import com.github.poi.xly.test.WorkbookTest;

public class XLYFormatterTest extends WorkbookTest {

    public final static String CUSTOM_BLUE = "051039";
    public final static String CUSTOM_GREEN = "a4bf3a";

    private XLYFormatter xlyFormatter;

    /**
     * Utility test method to check if colors and font have been correctly set
     * in excel.
     */
    public static void assertColors(String expectedFillColorCode, String expectedFontColorCode, Workbook workbook,
            CellStyle cellStyle) {
        final XSSFColor expectedFillColor = XLYFormatter.toColor(expectedFillColorCode);
        assertEquals(expectedFillColor.getIndexed(), cellStyle.getFillBackgroundColor());
        assertEquals(expectedFillColor.getIndexed(), cellStyle.getFillForegroundColor());
        assertEquals(FillPatternType.SOLID_FOREGROUND, cellStyle.getFillPatternEnum());
        final XSSFColor expectedFontColor = XLYFormatter.toColor(expectedFontColorCode);
        final Font font = workbook.getFontAt(cellStyle.getFontIndex());
        assertEquals(expectedFontColor.getIndexed(), font.getColor());
    }

    @Before
    public void createFormatter() {
        xlyFormatter = new XLYFormatter(workbook);
    }

    /**
     * Test constructor init of class attributes.
     */
    @Test
    public void testXLYFormatter() {
        assertNotNull(xlyFormatter.getDefaultCellStyle());
        assertNotNull(xlyFormatter.getDataFormat());
    }

    /**
     * Test nominal case (foreground!=WHITE / font!=BLACK)
     */
    @Test
    public void testFormatHeader() throws NoSuchFieldException {
        final XLYSheet xlySheet = getXLYSheet();
        final XLYColumn xlyColumn = xlySheet.columns()[0];
        Cell cell = createCell();
        xlyFormatter.formatHeader(xlyColumn, cell);
        final CellStyle cellStyle = cell.getCellStyle();
        assertAlignment(cellStyle);
        assertColors(CUSTOM_BLUE, CUSTOM_GREEN, workbook, cellStyle);

    }

    /**
     * Test specific case: (foreground=WHITE / font=BLACK)
     */
    @Test
    public void testFormatHeader_blackAndWhite() throws NoSuchFieldException {
        final XLYSheet xlySheet = getXLYSheet();
        final XLYColumn xlyColumn = xlySheet.columns()[1];
        final Cell cell = createCell();
        xlyFormatter.formatHeader(xlyColumn, cell);
        final CellStyle cellStyle = cell.getCellStyle();
        assertAlignment(cellStyle);
        assertEquals(HSSFColor.WHITE.index, cellStyle.getFillForegroundColor());
        assertEquals(FillPatternType.NO_FILL, cellStyle.getFillPatternEnum());
        final Font font = workbook.getFontAt(cellStyle.getFontIndex());
        assertEquals(HSSFColor.BLACK.index, font.getColor());
    }

    @Test
    public void testFormatCell_null() throws NoSuchFieldException {
        final Cell cell = formatCell("aString", null);
        assertEquals("", cell.getStringCellValue());
        assertEquals(xlyFormatter.getDefaultCellStyle(), cell.getCellStyle());
    }

    @Test
    public void testFormatCell_string() throws NoSuchFieldException {
        final Cell cell = formatCell("aString", "aValue");
        assertEquals(CellType.STRING, cell.getCellTypeEnum());
        assertEquals("aValue", cell.getStringCellValue());
        assertEquals(xlyFormatter.getDefaultCellStyle(), cell.getCellStyle());
    }

    @Test
    public void testFormatCell_enum() throws NoSuchFieldException {
        final Cell cell = formatCell("anEnum", DayOfWeek.FRIDAY);
        assertEquals(CellType.STRING, cell.getCellTypeEnum());
        assertEquals("FRIDAY", cell.getStringCellValue());
        assertEquals(xlyFormatter.getDefaultCellStyle(), cell.getCellStyle());
    }

    @Test
    public void testFormatCell_boolean() throws NoSuchFieldException {
        final Cell cell = formatCell("aBoolean", true);
        assertEquals(CellType.BOOLEAN, cell.getCellTypeEnum());
        assertEquals(true, cell.getBooleanCellValue());
        assertEquals(xlyFormatter.getDefaultCellStyle(), cell.getCellStyle());
    }

    @Test
    public void testFormatCell_double() throws NoSuchFieldException {
        final Cell cell = formatCell("aDouble", 123d);
        assertEquals(CellType.NUMERIC, cell.getCellTypeEnum());
        assertEquals(123.0d, cell.getNumericCellValue(), 0.1);
        assertEquals(xlyFormatter.getDefaultCellStyle(), cell.getCellStyle());
    }

    @Test
    public void testFormatCell_integer() throws NoSuchFieldException {
        final Cell cell = formatCell("anInteger", 456);
        assertEquals(CellType.NUMERIC, cell.getCellTypeEnum());
        assertEquals(456.0d, cell.getNumericCellValue(), 0.1);
        assertEquals(xlyFormatter.getDefaultCellStyle(), cell.getCellStyle());
    }

    @Test
    public void testFormatCell_long() throws NoSuchFieldException {
        final Cell cell = formatCell("aLong", 789l);
        assertEquals(CellType.NUMERIC, cell.getCellTypeEnum());
        assertEquals(789.0, cell.getNumericCellValue(), 0.1);
        assertEquals(xlyFormatter.getDefaultCellStyle(), cell.getCellStyle());
    }

    @Test
    public void testFormatCell_date() throws NoSuchFieldException {
        final LocalDate localDate = LocalDate.of(2018, 12, 28);
        final Date date = new Date(localDate.toEpochDay());
        final Cell cell = formatCell("aDate", date);
        assertEquals(CellType.NUMERIC, cell.getCellTypeEnum());
        assertTrue(cell.getDateCellValue().equals(date));
        final CellStyle cellStyle = cell.getCellStyle();
        assertEquals("dd/MM/YYYY", cellStyle.getDataFormatString());
    }

    @Test
    public void testAutoSizing() {
        final SXSSFSheet sheet = mock(SXSSFSheet.class);
        final int nbCols = 2;
        xlyFormatter.autoSizing(sheet, nbCols);
        verify(sheet).trackAllColumnsForAutoSizing();
        verify(sheet, times(nbCols)).autoSizeColumn(anyInt());
    }

    /**
     * Test for {@link XLYFormatter#addErrorMessage(Cell, String)}
     */
    @Test
    public void testAddErrorMessage_cellLevel() {
        final Cell cell = createErrorSheet();
        final String comment = "a cell level error message";
        xlyFormatter.addErrorMessage(cell, comment);
        assertCellsWithErrorStyle(workbook, 0, "A2");
        assertCellsWithErrorStyle(workbook, 0, "B2");
        final Row row = cell.getRow();
        final short lastCellNum = row.getLastCellNum();
        assertEquals(comment, row.getCell(lastCellNum - 1).getStringCellValue());
    }

    /**
     * Append the new error mesage to an existing error message. <br/>
     * 
     * Test for {@link XLYFormatter#addErrorMessage(Cell, String)}
     */
    @Test
    public void testAddErrorMessage_cellLevel_withExistingError() {
        final Cell cell = createErrorSheet();
        final String comment = "a cell level error message";
        xlyFormatter.addErrorMessage(cell, comment);
        final String anotherMessage = "another comment";
        xlyFormatter.addErrorMessage(cell, anotherMessage);
        assertCellsWithErrorStyle(workbook, 0, "A2");
        assertCellsWithErrorStyle(workbook, 0, "B2");
        final Row row = cell.getRow();
        final short lastCellNum = row.getLastCellNum();
        final String expected = String.format("%s\n%s", comment, anotherMessage);
        assertEquals(expected, row.getCell(lastCellNum - 1).getStringCellValue());
    }

    /**
     * Test for {@link XLYFormatter#addErrorMessage(Row, String)}
     */
    @Test
    public void testAddErrorMessage_rowLevel() {
        final Cell cell = createErrorSheet();
        final Row row = cell.getRow();
        final String comment = "a row level error message";
        xlyFormatter.addErrorMessage(row, comment);
        assertCellsWithErrorStyle(workbook, 0, "B2");
        final short lastCellNum = row.getLastCellNum();
        assertEquals(comment, row.getCell(lastCellNum - 1).getStringCellValue());
    }

    /**
     * Check that we don't add a comment to a row with no cell at all.<br/>
     * Test for {@link XLYFormatter#addErrorMessage(Row, String)}
     */
    @Test
    public void testAddErrorMessage_rowLevel_noCells() {
        final int rownum = 0;
        final Row row = workbook.createSheet("testSheet").createRow(rownum);
        // don't create any cell on purpose
        xlyFormatter.addErrorMessage(row, "a row level error message");
        assertEquals(-1, row.getLastCellNum());
    }

    private Cell formatCell(String fieldName, Object value) throws NoSuchFieldException {
        final Field field = getField(fieldName);
        final XLYColumn xlyColumn = getXLYColumn(fieldName);
        final Cell cell = createCell();
        xlyFormatter.formatCell(field, xlyColumn, cell, value);
        return cell;
    }

    private Cell createCell() {
        return workbook.createSheet("testSheet").createRow(0).createCell(0);
    }

    private void assertAlignment(CellStyle cellStyle) {
        assertEquals(VerticalAlignment.CENTER, cellStyle.getVerticalAlignmentEnum());
        assertEquals(HorizontalAlignment.CENTER, cellStyle.getAlignmentEnum());
    }

    private XLYSheet getXLYSheet() throws NoSuchFieldException {
        return MyWorkbook.class.getDeclaredField("rows").getAnnotation(XLYSheet.class);
    }

    private Field getField(String field) throws NoSuchFieldException {
        return MyRow.class.getDeclaredField(field);
    }

    private XLYColumn getXLYColumn(String fieldName) throws NoSuchFieldException {
        final XLYColumn[] columns = getXLYSheet().columns();
        return Arrays.stream(columns).filter(col -> col.field().equals(fieldName)).findFirst().get();
    }

    /**
     * Create a sheet with a heade and a cell value.
     * 
     * @return the cell created in this sheet with the value
     */
    private Cell createErrorSheet() {
        final XSSFSheet sheet = workbook.createSheet("test");
        final Cell header = sheet.createRow(0).createCell(0);
        header.setCellValue("a header");
        final Cell cell = sheet.createRow(1).createCell(0);
        cell.setCellValue("a value");
        return cell;
    }

    @XLYWorkbook
    public class MyWorkbook {
        @XLYSheet(name = "Scenario", type = MyRow.class, columns = {
                @XLYColumn(field = "aString", headerTitle = "A String", headerForeground = CUSTOM_BLUE, headerFont = CUSTOM_GREEN),
                @XLYColumn(field = "aDate", datePattern = "dd/MM/YYYY", headerTitle = "A Date", headerForeground = Colors.WHITE, headerFont = Colors.BLACK),
                @XLYColumn(field = "anInteger", headerTitle = "An Integer"),
                @XLYColumn(field = "aDouble", headerTitle = "A Double"),
                @XLYColumn(field = "aLong", headerTitle = "A Long"),
                @XLYColumn(field = "aBoolean", headerTitle = "A Boolean"),
                @XLYColumn(field = "anEnum", headerTitle = "An Enum") })
        private List<MyRow> rows;

        public List<MyRow> getRows() {
            return rows;
        }

        public void setRows(List<MyRow> rows) {
            this.rows = rows;
        }
    }

    /**
     * Set class attributs to public to avoid getter/setter boilerplate code.
     */
    public class MyRow {
        public String aString;
        public Date aDate;
        public Integer anInteger;
        public Double aDouble;
        public Long aLong;
        public Boolean aBoolean;
        public DayOfWeek anEnum;
    }
}
