package com.github.poi.xly.validation;

import java.lang.reflect.Field;

import org.apache.poi.ss.usermodel.Cell;

import com.github.poi.xly.annotation.XLYColumn;

/**
 * A simple pojo object to hold values necessary for {@link CellValidator}.
 */
public class CellContext {

    private Cell cell;

    private String cellValue;

    private Field field;

    private XLYColumn xlyColumn;

    public CellContext() {
    }

    public CellContext(Cell cell, String cellValue, XLYColumn xlyColumn, Field field) {
        this.cell = cell;
        this.cellValue = cellValue;
        this.xlyColumn = xlyColumn;
        this.field = field;
    }

    public Cell getCell() {
        return cell;
    }

    public String getCellValue() {
        return cellValue;
    }

    public Field getField() {
        return field;
    }

    public XLYColumn getXlyColumn() {
        return xlyColumn;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }

    public void setCellValue(String cellValue) {
        this.cellValue = cellValue;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public void setXlyColumn(XLYColumn xlyColumn) {
        this.xlyColumn = xlyColumn;
    }
}
