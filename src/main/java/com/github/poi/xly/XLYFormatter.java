package com.github.poi.xly;

import java.awt.Color;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.Date;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.poi.xly.annotation.XLYColumn;

/**
 * Hold various formatting and styling methods used by {@link XLYExporter}
 * and{@link XLYValidator}.
 */
public class XLYFormatter {

    private static Logger logger = LoggerFactory.getLogger(XLYFormatter.class);

    public static final XSSFColor RED = new XSSFColor(Color.red);

    public static final short RED_INDEX = HSSFColor.RED.index;

    public static XSSFColor toColor(String hexacode) {
        int red = Integer.valueOf(hexacode.substring(0, 2), 16);
        int green = Integer.valueOf(hexacode.substring(2, 4), 16);
        int blue = Integer.valueOf(hexacode.substring(4, 6), 16);
        final XSSFColor xssfColor = new XSSFColor(new Color(red, green, blue));
        return xssfColor;
    }

    /** default cell formating for data */
    private final DataFormat dataFormat;

    /** default cell formating */
    private final CellStyle defaultCellStyle;

    /** error cell formating */
    private final CellStyle errorStyle;

    private final Workbook workbook;

    public XLYFormatter(Workbook workbook) {
        this.workbook = workbook;
        defaultCellStyle = workbook.createCellStyle();
        dataFormat = workbook.createDataFormat();
        errorStyle = initErrorStyle();
    }

    private void addCommentCell(String message, final int commentIndex, final Row row) {
        final Cell commentCell = row.getCell(commentIndex);
        if (commentCell == null) {
            // create the cell
            final Cell comment = row.createCell(commentIndex);
            comment.setCellStyle(errorStyle);
            comment.setCellValue(message);
        } else {
            // append value to existing
            final Cell comment = commentCell;
            final StringBuilder newValue = new StringBuilder(comment.getStringCellValue());
            newValue.append("\n");
            newValue.append(message);
            comment.setCellValue(newValue.toString());
        }
    }

    /**
     * Set erorStyle to cell + add error message at the end of the row. <br/>
     * Use in case of cell validation error.
     */
    public void addErrorMessage(Cell cell, String message) {
        final Row row = cell.getRow();
        final int commentIndex = getErrorCommentIndex(row);
        cell.setCellStyle(errorStyle);
        // add a 'comment' cell to end of row for each cell in error
        addCommentCell(message, commentIndex, row);
    }

    /**
     * Add error message at the end of the row. <br/>
     * Use in case of row validation error.
     */
    public void addErrorMessage(Row row, String message) {
        final int commentIndex = getErrorCommentIndex(row);
        if (commentIndex < 0) {
            final String msg = "Unable to add error comment to current line. The current row {} of sheet {} doesn't seems to have any cell (lastCellNum {}).";
            logger.error(msg, row.getRowNum(), row.getSheet().getSheetName(), commentIndex);
            return;
        }
        addCommentCell(message, commentIndex, row);
    }

    /**
     * Resize column based on content.
     */
    public void autoSizing(SXSSFSheet sheet, int length) {
        sheet.trackAllColumnsForAutoSizing();
        for (int k = 0; k < length; k++) {
            sheet.autoSizeColumn(k);
        }
    }

    public void formatCell(Field field, XLYColumn xlyColumn, Cell cell, Object value) {
        if (value != null) {
            if (field.getType().isAssignableFrom(Boolean.class)) {
                cell.setCellValue(value.toString());
                cell.setCellType(CellType.BOOLEAN);
            } else if (isNumeric(field)) {
                cell.setCellValue(Double.valueOf(value.toString()));
                cell.setCellType(CellType.NUMERIC);
            } else if (isDate(field)) {
                if (field.getType().isAssignableFrom(Date.class)) {
                    cell.setCellValue((Date) value);
                } else {
                    cell.setCellValue(value.toString());
                }
            } else {
                cell.setCellValue(value.toString());
                cell.setCellType(CellType.STRING);
            }

            // set cell style
            if (isDate(field)) {
                final CellStyle dateCellStyle = workbook.createCellStyle();
                dateCellStyle.setDataFormat(dataFormat.getFormat(xlyColumn.datePattern()));
                cell.setCellStyle(dateCellStyle);
            } else {
                cell.setCellStyle(defaultCellStyle);
            }
        }
    }

    public void formatHeader(final XLYColumn xlyColumn, final Cell cell) {
        final XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        // change color foreground
        if (!Colors.WHITE.equals(xlyColumn.headerForeground())) {
            // convert HEX format color (see XLYPalette.enum) to RGB format
            // color
            final XSSFColor xssfColor = toColor(xlyColumn.headerForeground());
            style.setFillBackgroundColor(xssfColor);
            style.setFillForegroundColor(xssfColor);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        } else {
            // default foreground color
            style.setFillForegroundColor(HSSFColor.WHITE.index);
            style.setFillPattern(FillPatternType.NO_FILL);
        }
        // change color font
        final Font font = workbook.createFont();
        if (!Colors.BLACK.equals(xlyColumn.headerFont())) {
            final int fontColor = Integer.parseInt(xlyColumn.headerFont(), 16);
            final XSSFColor xssfColor = new XSSFColor(new Color(fontColor));
            ((XSSFFont) font).setColor(xssfColor);
        } else { // default font color
            font.setColor(HSSFColor.BLACK.index);
        }
        style.setFont(font);
        cell.setCellStyle(style);
    }

    public DataFormat getDataFormat() {
        return dataFormat;
    }

    public CellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

    /**
     * @return headers last column + 1
     */
    private int getErrorCommentIndex(Row row) {
        final Row headers = row.getSheet().getRow(0);
        final int commentIndex = headers.getLastCellNum(); // see javadoc
                                                           // returns PLUS ONE
        return commentIndex;
    }

    /**
     * Create the style for further usage.
     */
    private CellStyle initErrorStyle() {
        final CellStyle style = workbook.createCellStyle();
        // add foreground red
        style.setFillForegroundColor(RED_INDEX);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillBackgroundColor(RED_INDEX);
        // add font white
        final Font font = workbook.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        return style;
    }

    private boolean isDate(Field field) {
        return field.getType().isAssignableFrom(Date.class) || field.getType().isAssignableFrom(LocalDate.class);
    }

    private boolean isNumeric(Field field) {
        return field.getType().isAssignableFrom(Double.class) || field.getType().isAssignableFrom(Integer.class)
                || field.getType().isAssignableFrom(Long.class);
    }
}
