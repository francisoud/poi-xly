package com.github.poi.xly;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.poi.xly.XLYException.XLYError;
import com.github.poi.xly.annotation.XLYColumn;
import com.github.poi.xly.annotation.XLYSheet;
import com.github.poi.xly.validation.CellValidatorManager;
import com.github.poi.xly.validation.ConstraintLocator;
import com.github.poi.xly.validation.RowValidatorManager;

/**
 * Validate the content of the excel file using the validators define
 * the @Workbook.
 */
public class XLYValidator {

    private static final Logger logger = LoggerFactory.getLogger(XLYValidator.class);

    private CellValidatorManager cellValidatorManager;

    private final ConstraintLocator constraintLocator;

    private RowValidatorManager rowValidatorManager;

    private Class<?> workbookClass;

    private XLYMetadataParser xlyMetadataParser;

    /**
     * 
     * @param workbookClass
     *            the class annotated with @XLYWorbook representing the
     *            inputStream to validate
     * @param constraintLocator
     */
    public XLYValidator(ConstraintLocator constraintLocator) {
        if (constraintLocator == null) {
            throw new IllegalArgumentException("constraintLocator must not be null");
        }
        this.constraintLocator = constraintLocator;
    }

    public Class<?> getWorkbookClass() {
        return workbookClass;
    }

    /**
     * Importing an excel in protected mode make poi crash with "Zip bomb
     * detected" unrelated message :(.<br/>
     * Refer to ZipSecureFile.java line:257
     * 
     * @see ZipSecureFile
     */
    private void handleIOException(IOException e) {
        if (e.getCause() != null && e.getCause().getMessage() != null
                && e.getCause().getMessage().startsWith("Zip bomb detected!")) {
            throw new XLYException(XLYError.PROTECTED_VIEW_ENABLE, e);
        }
        throw new RuntimeException(e.getMessage(), e);
    }

    /**
     * Case where the sheetName is define in the code but not available in the
     * imported workbook.
     */
    private void handleUnexistingSheet(XSSFWorkbook workbook, final XLYSheet xlySheet) {
        final StringBuilder sb = new StringBuilder();
        final Iterator<Sheet> sheetIterator = workbook.iterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            sb.append(sheet.getSheetName());
            sb.append(" ");
        }
        logger.error("Unable to find sheet with name: {} during excel import (excel sheetNames: {});", xlySheet.name(),
                sb);
        throw new XLYException(XLYError.MISSING_SHEET);
    }

    public boolean isValid(final InputStream inputStream, OutputStream outputStream) {
        final Set<String> messages = validate(inputStream, outputStream);
        if (!messages.isEmpty()) {
            logger.info("Excel validation errors: {}", String.join(",", messages));
        }
        return messages.isEmpty();
    }

    public void setWorkbookClass(Class<?> workbookClass) {
        this.workbookClass = workbookClass;
    }

    /**
     * @param inputStream
     *            the .xlsx file
     * @param outputStream
     *            the modified input .xlsx file + errors (a.k.a excel red cells
     *            + message)
     * @return an empty set if no error otherwise a list of error messages
     */
    public Set<String> validate(InputStream inputStream, OutputStream outputStream) {
        if (workbookClass == null || constraintLocator == null) {
            throw new IllegalArgumentException("workbookClass must not be null");
        }
        // Store the unique list of error messages.
        // Using a {@link Set} instead of a {@link List} helps us remove
        // duplicate error messages.
        // Reset to empty set when calling {@link #validate(InputStream,
        // OutputStream)}
        final Set<String> violations = new HashSet<>();
        xlyMetadataParser = new XLYMetadataParser(workbookClass);
        XSSFWorkbook workbook = null;
        try {
            // XSSFWorkbook(inputStream) load all the excel file in memory but
            // using the low memory footprint version would create complicated
            // code
            // see: http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api
            workbook = new XSSFWorkbook(inputStream);
            final XLYFormatter xlyFormatter = new XLYFormatter(workbook);
            cellValidatorManager = new CellValidatorManager(constraintLocator, xlyFormatter);
            rowValidatorManager = new RowValidatorManager(constraintLocator, xlyFormatter);
            final List<XLYSheet> xlySheets = xlyMetadataParser.getSheets();
            for (final XLYSheet xlySheet : xlySheets) {
                if (xlySheet.toImport()) {
                    final XSSFSheet sheet = workbook.getSheet(xlySheet.name());
                    if (sheet == null) {
                        handleUnexistingSheet(workbook, xlySheet);
                    }
                    violations.addAll(validateSheet(sheet, xlySheet));
                }
            }
            if (!violations.isEmpty()) {
                workbook.write(outputStream);
            }
        } catch (IOException e) {
            handleIOException(e);
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                logger.error(e.getMessage(), e);
            }
        }
        return violations;
    }

    /**
     * For each sheet rows run cells validations then row validations. <br/>
     * <b>Gotcha</b>: if row.getCell() is null we create the cell so that we can
     * add a error color to it later with
     * {@link XLYFormatter#addErrorMessage(Cell, String)}
     * 
     * @return
     */
    private Set<String> validateSheet(XSSFSheet sheet, XLYSheet xlySheet) {
        final Set<String> sheetViolations = new HashSet<>();
        final Iterator<Row> rowIterator = sheet.iterator();
        Row row = rowIterator.next(); // skip first row (a.k.a headers)
        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            final XLYColumn[] xlyColumns = xlySheet.columns();
            for (int column = 0; column < xlyColumns.length; column++) {
                final XLYColumn xlyColumn = xlyColumns[column];
                Cell cell = row.getCell(column);
                if (cell == null) {
                    cell = row.createCell(column); // see method javadoc
                }
                final Field field = xlyMetadataParser.getField(xlySheet);
                sheetViolations.addAll(cellValidatorManager.validate(cell, xlyColumn, field));
            }
        }
        if (sheetViolations.isEmpty()) {
            // do the rows validation only if no cell errors
            // avoid iterating over all rows one more time, anyway this sheet is
            // already not valid
            // the end user will on start to see those errors once he has fix
            // cells errors
            // deliberately choose performance improvement over user experience
            sheetViolations.addAll(rowValidatorManager.validateRows(sheet, xlySheet));
        }
        if (!sheetViolations.isEmpty()) {
            sheet.setTabColor(XLYFormatter.RED);
        }
        return sheetViolations;
    }

}
