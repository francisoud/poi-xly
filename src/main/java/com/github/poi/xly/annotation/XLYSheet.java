package com.github.poi.xly.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.github.poi.xly.validation.NoContraint;
import com.github.poi.xly.validation.RowConstraint;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface XLYSheet {

    public XLYColumn[] columns();

    /**
     * Name of the sheet
     */
    public String name();

    public Class<? extends RowConstraint>[] rowValidator() default NoContraint.class;

    /**
     * if set to false the sheet will be ignored during import process
     */
    public boolean toImport() default true;

    /**
     * Type of the inside the List&lt;type&gt; in the workbook.
     */
    public Class<?> type();
}
