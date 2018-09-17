package com.github.poi.xly.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.github.poi.xly.validation.Constraint;
import com.github.poi.xly.validation.NoContraint;

@Target(ElementType.FIELD)
@Retention(value = RetentionPolicy.RUNTIME)
public @interface XLYColumn {

    static String EMPTY_DATE_PATTERN = "";

    public Class<? extends Constraint> cellValidator() default NoContraint.class;

    /**
     * Date pattern of the field (if the field is a date). <br/>
     * Only use for export (not for validation)
     */
    public String datePattern() default EMPTY_DATE_PATTERN;

    /**
     * The attribut name of the java POJO associated with this column.<br/>
     * Example: <code>@XLYColumn(field="creationDate")<code> where
     * 
     * <pre>
     * class Flown {
     *     private Date creationDate;
     * }
     * </pre>
     */
    public String field();

    /**
     * font color in hexadecimal
     */
    public String headerFont() default "000000";

    /**
     * foreground color in hexadecimal
     */
    public String headerForeground() default "ffffff";

    /**
     * column name
     */
    public String headerTitle();

    /**
     * cell content mandatory
     */
    public boolean mandatory() default false;

    /**
     * Regex validator
     */
    public String pattern() default "";
}
