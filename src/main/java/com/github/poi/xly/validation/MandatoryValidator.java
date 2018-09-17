package com.github.poi.xly.validation;

public class MandatoryValidator implements CellValidator {

    @Override
    public String validate(CellContext cellContext) {
        if (cellContext.getXlyColumn().mandatory() == false) {
            return null;
        }
        final String cellValue = cellContext.getCellValue();
        if (cellValue != null) {
            if (!cellValue.isEmpty()) {
                return null;
            }
        }
        return "Field required";
    }
}
