package com.github.poi.xly;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.ConvertUtilsBean;
import org.apache.commons.beanutils.ConvertUtilsBean2;
import org.apache.commons.beanutils.Converter;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.poi.xly.annotation.XLYColumn;
import com.github.poi.xly.annotation.XLYSheet;

/**
 * Construct the object and sub-objects to be saved to the database.
 */
public class XLYImporter<T> {

    private static final Logger logger = LoggerFactory.getLogger(XLYImporter.class);

    private ConvertUtilsBean converter = new ConvertUtilsBean2();

    private Class<T> workbookClass;

    private XLYMetadataParser xlyMetadataParser;

    private List<?> createObjects(XSSFSheet sheet, XLYSheet xlySheet) {
        final Map<String, Class<?>> columnsTypes = xlyMetadataParser.getColumnTypes(xlySheet);
        final List<Object> beans = new ArrayList<>();
        final Iterator<Row> rowIterator = sheet.iterator();
        Row row = rowIterator.next(); // skip first row (a.k.a header)
        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            try {
                final Object bean = xlySheet.type().newInstance();
                final XLYColumn[] xlyColumns = xlySheet.columns();
                for (int i = 0; i < xlyColumns.length; i++) {
                    final Cell cell = row.getCell(i);
                    if (cell != null) {
                        final CellType cellType = cell.getCellTypeEnum();
                        final Object value = getCellValue(cell, cellType);
                        final Class<?> type = columnsTypes.get(xlyColumns[i].field());
                        final Object converted = converter.convert(value, type);
                        PropertyUtils.setProperty(bean, xlyColumns[i].field(), converted);
                    }
                }
                beans.add(bean);
            } catch (ReflectiveOperationException e) {
                logger.error(e.getMessage(), e);
            }
        }
        return beans;
    }

    private Object getCellValue(final Cell cell, final CellType cellType) {
        final Object value;
        switch (cellType) {
        case NUMERIC:
            final CellStyle style = cell.getCellStyle();
            final int formatNo = style.getDataFormat();
            final String formatString = style.getDataFormatString();
            if (DateUtil.isADateFormat(formatNo, formatString)) {
                value = cell.getDateCellValue();
            } else {
                value = cell.getNumericCellValue();
            }
            break;
        case BOOLEAN:
            value = cell.getBooleanCellValue();
            break;
        case STRING:
            value = cell.getStringCellValue();
            break;
        case BLANK:
            value = "";
            break;
        default:
            final String msg = String.format("Unhandle cell type: %s for cell: %s", cellType.name(), cell.getAddress());
            logger.error(msg);
            throw new IllegalStateException(msg);
        }
        return value;
    }

    public Class<T> getWorkbookClass() {
        return workbookClass;
    }

    public void register(Converter converter, Class<?> clazz) {
        this.converter.register(converter, clazz);
    }

    /**
     * @return the workbook object as define in
     *         {@link XLYImporter#XLYImporter(Class)}
     */
    public T save(InputStream inputStream) {
        if (inputStream == null || workbookClass == null) {
            throw new IllegalArgumentException("inputStream or workbookClass must not be null");
        }
        XSSFWorkbook workbook = null;
        try {
            // XSSFWorkbook(inputStream) load all the excel file in memory but
            // using the low memory footprint version would create complicated
            // code
            // see: http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api
            workbook = new XSSFWorkbook(inputStream);
            final T xlyWorkbook = workbookClass.newInstance();
            xlyMetadataParser = new XLYMetadataParser(xlyWorkbook);
            final List<XLYSheet> xlySheets = xlyMetadataParser.getSheets();
            for (final XLYSheet xlySheet : xlySheets) {
                if (xlySheet.toImport()) {
                    final XSSFSheet sheet = workbook.getSheet(xlySheet.name());
                    // the sheet is suppose to be there because it is tested
                    // during
                    // XLYValidation#handleUnexistingSheet()
                    final List<?> beans = createObjects(sheet, xlySheet);
                    final Field field = xlyMetadataParser.getField(xlySheet);
                    PropertyUtils.setProperty(xlyWorkbook, field.getName(), beans);
                }
            }
            return xlyWorkbook;
        } catch (final IOException | ReflectiveOperationException e) {
            logger.error(e.getMessage(), e);
            throw new XLYException(e);
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                logger.error(e.getMessage(), e);
            }
        }
    }

    public void setWorkbookClass(Class<T> workbookClass) {
        this.workbookClass = workbookClass;
    }
}
