package com.github.poi.xly.validation;

import java.util.ArrayList;
import java.util.List;

public class EnumValidator implements CellValidator {

    @Override
    public String validate(CellContext cellContext) {
        final String cellValue = cellContext.getCellValue();
        if (cellValue == null) {
            return null;
        }
        final Class<?> type = cellContext.getField().getType();
        if (!type.isEnum()) {
            return null;
        }
        final Object[] enumValues = type.getEnumConstants();
        final List<String> enumsList = new ArrayList<>();
        for (final Object enumValue : enumValues) {
            if (enumValue.toString().equals(cellValue)) {
                return null; // found the value in the enum list
            }
            enumsList.add(enumValue.toString());
        }
        final String possiblesValues = String.join(" or ", enumsList);
        final String msg = String.format("%s not exists. Did you mean : %s ?", cellValue, possiblesValues);
        return msg;
    }
}
