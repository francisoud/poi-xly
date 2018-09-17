package com.github.poi.xly.validation;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

/**
 * Check if at least one row is duplicated. <br/>
 * A line is considered as duplicate if all values in columns define in
 * {@link DuplicatedRowConstraint#columnsHeaders()} are identical.
 */
public abstract class DuplicatedRowConstraint implements RowConstraint {

    public static final String DUPLICATE_ERROR = "Duplicate line";
    private static final char SEP = ';';

    private String buildUniqueKey(final Set<Integer> cellNums, final DataFormatter dataFormatter, final Row row) {
        final StringBuilder uniqueKey = new StringBuilder();
        for (Integer cellnum : cellNums) {
            final Cell cell = row.getCell(cellnum);
            uniqueKey.append(dataFormatter.formatCellValue(cell));
            uniqueKey.append(SEP);
        }
        return uniqueKey.toString();
    }

    /**
     * Extract column num to be checked for the provided columns header.<br/>
     * <b>Warning</b> move the iterator of one row.
     */
    private Set<Integer> getCellNums(Iterator<Row> rowIterator, final Set<String> fields) {
        final Set<Integer> cellNums = new HashSet<>();
        final Row headers = rowIterator.next();
        final Iterator<Cell> headerIterator = headers.cellIterator();
        while (headerIterator.hasNext()) {
            final Cell header = headerIterator.next();
            final String title = header.getStringCellValue();
            if (fields.contains(title)) {
                cellNums.add(header.getColumnIndex());
            }
        }
        if (fields.size() != cellNums.size()) {
            handleHeadersMismatch(fields, headers, cellNums);
        }
        return cellNums;
    }

    /**
     * Compare expected fields vs actual columns headers and throw an exception
     * if mismatch.
     */
    private void handleHeadersMismatch(final Set<String> fields, final Row headers, Set<Integer> cellNums) {
        final Stream<Cell> targetStream = StreamSupport
                .stream(Spliterators.spliteratorUnknownSize(headers.cellIterator(), Spliterator.ORDERED), false);
        final String headersTitles = targetStream.map(Cell::getStringCellValue).collect(Collectors.joining(","));
        final String msg = String.format("Expected %s columns header (%s), found: %s, all headers in excel: (%s)",
                fields.size(), String.join(",", fields), cellNums.size(), String.join(",", headersTitles));
        throw new IllegalStateException(msg);
    }

    /**
     * @see RowConstraint#validate(Iterator)
     */
    @Override
    public Map<Integer, String> validate(Iterator<Row> rowIterator) {
        final Map<Integer, String> violations = new HashMap<>();
        final Set<String> fields = columnsHeaders();
        final Set<Integer> cellNums = getCellNums(rowIterator, fields);
        final Set<String> uniqueKeys = new HashSet<>();
        final DataFormatter dataFormatter = new DataFormatter();
        while (rowIterator.hasNext()) {
            final Row row = rowIterator.next();
            final String uniqueKey = buildUniqueKey(cellNums, dataFormatter, row);
            if (uniqueKeys.contains(uniqueKey)) {
                violations.put(row.getRowNum(), DUPLICATE_ERROR);
                break; // stop on first duplicated line to avoid creating too
                       // many object in memory
            }
            uniqueKeys.add(uniqueKey.toString());
        }
        return violations;
    }
}
