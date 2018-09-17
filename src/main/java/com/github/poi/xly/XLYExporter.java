package com.github.poi.xly;

import static com.github.poi.xly.XLYException.XLYError.UNEXCEPTED_ERROR;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.poi.xly.annotation.XLYColumn;
import com.github.poi.xly.annotation.XLYSheet;

public class XLYExporter {

    private static final Logger logger = LoggerFactory.getLogger(XLYExporter.class);

    private SXSSFWorkbook workbook;

    private XLYFormatter xlyFormatter;

    private XLYMetadataParser xlyMetadataParser;

    private final Object xlyWorkbook;

    /**
     * @param workbook
     *            object annotated with @XLYWorkbook
     */
    public XLYExporter(Object xlyWorkbook) {
        this.xlyWorkbook = xlyWorkbook;
    }

    /**
     * 
     * Generate an excel file. <br/>
     * 
     * @param outputStream
     *            the outputstream to write the result to.
     * @return isSuccessful: true if generation without errors
     */
    public void export(OutputStream outputStream) {
        xlyMetadataParser = new XLYMetadataParser(xlyWorkbook);
        if (xlyWorkbook == null) {
            throw new IllegalArgumentException("workbook can't be null");
        }
        try {
            workbook = new SXSSFWorkbook();
            xlyFormatter = new XLYFormatter(workbook);
            final List<XLYSheet> xlySheets = xlyMetadataParser.getSheets();
            for (final XLYSheet xlySheet : xlySheets) {
                final SXSSFSheet sheet = createSheet(xlySheet);
                final List<?> beans = xlyMetadataParser.getValues(xlyWorkbook, xlySheet);
                populateSheet(sheet, xlySheet, beans);
            }
            workbook.write(outputStream);
        } catch (final IOException e) {
            logger.error(e.getMessage(), e);
            throw new XLYException(UNEXCEPTED_ERROR, e);
        } finally {
            if (workbook != null) {
                workbook.dispose();
            }
        }
    }

    /**
     * Create a sheet and set columns headers.
     */
    private SXSSFSheet createSheet(final XLYSheet xlySheet) {
        final SXSSFSheet sheet = workbook.createSheet(xlySheet.name());
        final SXSSFRow headerRow = sheet.createRow(0);
        final XLYColumn[] xlyColumns = xlySheet.columns();
        for (int column = 0; column < xlyColumns.length; column++) {
            final XLYColumn xlyColumn = xlyColumns[column];
            final SXSSFCell cell = headerRow.createCell(column);
            cell.setCellValue(xlyColumn.headerTitle());
            xlyFormatter.formatHeader(xlyColumn, cell);
        }
        xlyFormatter.autoSizing(sheet, xlyColumns.length);
        return sheet;
    }

    private Object getProperty(Object bean, XLYColumn xlyColumn) {
        try {
            return PropertyUtils.getProperty(bean, xlyColumn.field());
        } catch (ReflectiveOperationException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e.getMessage(), e);
        }
    }

    private void populateRow(SXSSFRow row, XLYColumn[] xlyColumns, Map<String, Field> fields, Object bean) {
        for (int column = 0; column < xlyColumns.length; column++) {
            final XLYColumn xlyColumn = xlyColumns[column];
            final Field field = fields.get(xlyColumn.field());
            final SXSSFCell cell = row.createCell(column);
            final Object value = getProperty(bean, xlyColumn);
            xlyFormatter.formatCell(field, xlyColumn, cell, value);
        }
    }

    private void populateSheet(final SXSSFSheet sheet, final XLYSheet xlySheet, List<?> beans) {
        int rownum = 1; // 1 because we skip the first row (a.k.a header)
        final Map<String, Field> fields = xlyMetadataParser.getDataFields(xlySheet.type(), xlySheet.columns());
        for (final Object bean : beans) {
            final SXSSFRow row = sheet.createRow(rownum);
            populateRow(row, xlySheet.columns(), fields, bean);
            rownum++;
        }
    }
}
