package com.kyle.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Ignore all unannotated fields.
 *
 * @package: com.kyle.excel.annotation
 * @className: ExcelIgnoreUnannotated
 * @author: Kyle.Y.Li
 * @since 1.0.0 2020-04-4/29/2020 13:17
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelIgnoreUnannotated {}