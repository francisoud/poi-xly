package com.github.poi.xly;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.poi.xly.annotation.XLYColumn;
import com.github.poi.xly.annotation.XLYSheet;
import com.github.poi.xly.annotation.XLYWorkbook;

public class XLYMetadataParser {

    private static final Logger logger = LoggerFactory.getLogger(XLYMetadataParser.class);

    private final LinkedHashMap<XLYSheet, Field> sheets = new LinkedHashMap<>();

    private final Class<?> workbookClass;

    /**
     * @param the
     *            expect spreadsheet class annotated with @XLYWorkbook
     */
    public XLYMetadataParser(Class<?> workbookClass) {
        if (workbookClass == null) {
            throw new IllegalArgumentException("workbookClass can't be null");
        }
        this.workbookClass = workbookClass;
        assertWorkbookAnnotation(workbookClass);
        parseSheets();
    }

    /**
     * @param spreadsheet
     *            an object annotated with @XLYWorkbook
     */
    public XLYMetadataParser(Object workbook) {
        if (workbook == null) {
            throw new IllegalArgumentException("workbook can't be null");
        }
        workbookClass = workbook.getClass();
        assertWorkbookAnnotation(workbookClass);
        parseSheets();
    }

    /**
     * Check that object is correctly annotated.
     */
    private void assertWorkbookAnnotation(Class<?> workbookClass) {
        if (workbookClass.getAnnotation(XLYWorkbook.class) == null) {
            throw new IllegalArgumentException("workbookClass must be annotated with @XLYWorkbook");
        }
    }

    /**
     * @return a map of fieldName and their associated type (a.k.a
     *         Intger,Double, String)
     */
    public Map<String, Class<?>> getColumnTypes(XLYSheet xlySheet) {
        try {
            final Map<String, Class<?>> columnTypes = new HashMap<>();
            for (XLYColumn xlyColumn : xlySheet.columns()) {
                final Class<?> beanClass = xlySheet.type();
                final String fieldName = xlyColumn.field();
                final Field field = beanClass.getDeclaredField(fieldName);
                columnTypes.put(fieldName, field.getType());
            }
            return columnTypes;
        } catch (ReflectiveOperationException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e.getMessage(), e);
        }
    }

    /**
     * Create a Map<fieldName, Field> for each bean param define in
     * XLYColumn[]#type() <br/>
     * Avoid calling java.lang.reflect foreach(rows).
     */
    public Map<String, Field> getDataFields(Class<?> beanClass, XLYColumn[] xlyColumns) {
        final Map<String, Field> fields = new HashMap<>(xlyColumns.length);
        for (XLYColumn xlyColumn : xlyColumns) {
            final Field field = getField(beanClass, xlyColumn);
            fields.put(xlyColumn.field(), field);
        }
        return fields;
    }

    private Field getField(Class<?> beanClass, XLYColumn xlyColumn) {
        try {
            return beanClass.getDeclaredField(xlyColumn.field());
        } catch (ReflectiveOperationException e) {
            logger.error(e.getMessage(), e);
            throw new XLYException(e);
        }
    }

    public Field getField(XLYSheet xlySheet) {
        return sheets.get(xlySheet);
    }

    private Object getProperty(Object bean, Field field) {
        try {
            return PropertyUtils.getProperty(bean, field.getName());
        } catch (ReflectiveOperationException e) {
            logger.error(e.getMessage(), e);
            throw new XLYException(e);
        }
    }

    /**
     * warning: not *just* a simple getter.
     */
    public List<XLYSheet> getSheets() {
        return new ArrayList<>(sheets.keySet());
    }

    /***
     * @param workbook
     *            an object annotated with @XLYWorkbook
     * @param xlySheet
     *            the associated sheet
     * @return the list of beans (db tuples)
     */
    public List<?> getValues(final Object workbook, XLYSheet xlySheet) {
        final Field field = getField(xlySheet);
        final Object scrollableResults = getProperty(workbook, field);
        if (scrollableResults instanceof List<?>) {
            return (List<?>) scrollableResults;
        }
        final String msg = "Expected " + List.class.getCanonicalName() + " got "
                + scrollableResults.getClass().getCanonicalName();
        throw new IllegalStateException(msg);
    }

    private void parseSheets() {
        final Field[] fields = workbookClass.getDeclaredFields();
        for (final Field field : fields) {
            final XLYSheet annotation = field.getAnnotation(XLYSheet.class);
            if (annotation != null) {
                sheets.put(annotation, field);
            }
        }
    }
}
