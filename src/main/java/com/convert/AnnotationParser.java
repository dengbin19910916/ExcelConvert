package com.convert;

/**
 * Created by Dengbin on 2015/10/6.
 */
public interface AnnotationParser {

    /**
     * 解析类型中所有的注解信息。
     *
     * @param clazz 需要解析注解信息的类型
     * @param <T> 带有注解的类型
     */
    <T> T parse(Class clazz);
}
