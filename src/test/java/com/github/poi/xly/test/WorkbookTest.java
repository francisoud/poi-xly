package com.github.poi.xly.test;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;

/**
 * Create a workbook and close is so that we can create a simple Cell.
 */
public abstract class WorkbookTest {

    protected XSSFWorkbook workbook;

    @Before
    public void setup() {
        workbook = new XSSFWorkbook();
    }

    @After
    public void tearDown() throws IOException {
        workbook.close();
    }

    protected Cell getCell(String value) {
        final Cell cell = workbook.createSheet("test").createRow(0).createCell(0);
        cell.setCellValue(value);
        return cell;
    }
}
