package com.github.poi.xly.validation;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.Before;
import org.junit.Test;

import com.github.poi.xly.test.WorkbookTest;

public class NoContraintTest extends WorkbookTest {

    private NoContraint noContraint;

    @Before
    public void createConstraint() {
        noContraint = new NoContraint();
    }

    @Test
    public void testValidate_cellNull() {
        final Cell cell = null;
        assertNull(noContraint.validate(cell));
    }

    @Test
    public void testValidate_cellEmpty() {
        final Cell cell = getCell("");
        assertNull(noContraint.validate(cell));
    }

    @Test
    public void testValidate_cell() {
        final Cell cell = getCell("a value");
        assertNull(noContraint.validate(cell));
    }

    @Test
    public void testValidate_rowIteratorNull() {
        final Iterator<Row> rowIterator = getRowIterator(null);
        noContraint.validate(rowIterator);
    }

    @Test
    public void testValidate_rowIteratorEmpty() {
        final Iterator<Row> rowIterator = getEmptyRowIterator();
        noContraint.validate(rowIterator);
    }

    @Test
    public void testValidate_rowIterator() {
        final Iterator<Row> rowIterator = getRowIterator("a value");
        noContraint.validate(rowIterator);
    }

    @Test
    public void testColumnsHeaders() {
        assertEquals(0, noContraint.columnsHeaders().size());
    }

    private Iterator<Row> getRowIterator(String cellValue) {
        final XSSFSheet sheet = workbook.createSheet();
        final Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(cellValue);
        return sheet.iterator();
    }

    private Iterator<Row> getEmptyRowIterator() {
        final XSSFSheet sheet = workbook.createSheet();
        return sheet.iterator();
    }
}
