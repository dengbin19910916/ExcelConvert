package com.convert.annotation;

import org.apache.poi.ss.usermodel.Cell;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.ElementType.METHOD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Excel文件的表头，也是数据列的列名。<br/>
 * 使用了此注解的属性必须拥有get/set方法。<br/>
 *
 * Created by Dengbin on 2015/10/4.
 */
@Target({METHOD, FIELD})
@Retention(RUNTIME)
@Documented
public @interface Title {
    String name() default "";

    /**
     * 单元格数据类型，默认为字符串类型。
     *
     * @see Cell#CELL_TYPE_BLANK
     * @see Cell#CELL_TYPE_NUMERIC
     * @see Cell#CELL_TYPE_STRING
     * @see Cell#CELL_TYPE_FORMULA
     * @see Cell#CELL_TYPE_BOOLEAN
     * @see Cell#CELL_TYPE_ERROR
     */
    int cellType() default Cell.CELL_TYPE_STRING;
}
