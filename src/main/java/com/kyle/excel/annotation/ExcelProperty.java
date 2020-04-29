package com.kyle.excel.annotation;

import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel数据Bean类的属性 和 excel列头 的映射关系 注解
 * <p>此注解只能标注在pojo的属性上，标注在其他地方不起作用。</p>
 * <p>标注此注解后对应的属性在在作为excel导出数据源时，会自动和excel的列头相映射。</p>
 *
 * @package: com.kyle.excel.annotation
 * @className: ExcelProperty
 * @author: Kyle.Y.Li
 * @since 1.0.0 2020-04-29 13:12
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {
    /**
     * 列头标题
     */
    String name() default "";

    /**
     * 列头位置(从0开始)
     */
    int index() default -1;

    /**
     * 列宽（不指定则自适应）
     */
    int with() default 0;

    /**
     * 前缀
     */
    String prefix() default "";

    /**
     * 后缀
     */
    String suffix() default "";

    /**
     * 时间格式
     */
    String dateFormat() default "MM/dd/yyyy HH:mm:ss";

    /**
     * 自定义Excel样式
     */
    String cellStyleJson() default "";

    /**
     * 自定义Excel样式
     */
    Class<? extends CellStyle> cellStyle() default CellStyle.class;
}