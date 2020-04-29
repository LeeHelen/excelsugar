package com.kyle.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Ignore convert excel
 *
 * @package: com.kyle.excel.annotation
 * @className: ExcelIgnore
 * @author: Kyle.Y.Li
 * @since 1.0.0 2020-04-4/29/2020 13:16
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelIgnore {}