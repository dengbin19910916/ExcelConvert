package com.convert;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;

/**
 * Excel文件读取，包括读取Excel表头，表内容和Excel表向JavaBean对象的转换。
 *
 * Created by Dengbin on 2015/10/3.
 */
public interface ExcelReader {

    /**
     * 读取Excel表格表头的内容。
     *
     * @param inputStream Excel文件的输入流
     * @return 表头内容的数组
     */
    String[] readHeader(InputStream inputStream);

    /**
     * 读取Excel表格表头的内容。
     *
     * @param workbook Excel工作簿对象
     * @return 表头内容的数组
     */
    String[] readHeader(Workbook workbook);

    /**
     * 读取Excel表格数据的内容。
     *
     * @param inputStream Excel文件的输入流
     * @param clazz 目标对象类型
     * @param <T> 目标对象参数
     * @return 目标对象数组
     */
    <T> T[] readContent(InputStream inputStream, Class<T> clazz) throws IOException, InvalidFormatException, InstantiationException, IllegalAccessException, InvocationTargetException;

    /**
     * 读取Excel表格数据的内容，将Excel中的数据转换成Javabean对象。
     *
     * @param workbook Excel工作簿对象
     * @param clazz 目标对象类型
     * @param <T> 目标对象参数
     * @return 目标对象数组
     */
    <T> T[] readContent(Workbook workbook, Class<T> clazz) throws IllegalAccessException, InstantiationException, InvocationTargetException, IOException, InvalidFormatException;
}
