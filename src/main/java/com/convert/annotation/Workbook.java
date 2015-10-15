package com.convert.annotation;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * 标识一个对象与Excel表格可以相互映射。
 * <p>
 *
 * @author Dengbin
 * Created by Dengbin on 2015/10/4.
 */
@Target({TYPE})
@Retention(RUNTIME)
@Documented
public @interface Workbook {

    /**
     * Excel文件名称，用于生成Excel文件。
     */
    String name();

    /**
     * Excel中的表单，可能有多张表单，
     * 默认为表单“Sheet1”。<br/>
     * 表单中的数据与Javabean对象进行相互转换。
     */
    Sheet[] sheets() default @Sheet;

    /**
     * Excel文件类型，默认为97-2003工作簿。
     *
     * @see Workbook.Type#XLS
     */
    Type type() default Type.XLS;

    /**
     * Excel文件类型
     *
     * @author Dengbin
     */
    enum Type {
        /**
         * Excel 97-2003 工作簿
         */
        XLS,
        /**
         * Excel 工作簿
         */
        XLSX,
        /**
         * Excel启用宏的工作簿
         */
        XLSM,
        /**
         * Excel 二进制工作簿
         */
        XLSB,
        /**
         * Excel 97-2003 模板
         */
        XLT,
        /**
         * Excel 模板
         */
        XLTX,
        /**
         * Excel 启用宏的模板
         */
        XLTM;

        /**
         * 返回Excel文件的类型。
         *
         * @param inputStream Excel文件的输入流
         * @return Excel文件的类型
         */
        public static Type getType(InputStream inputStream) throws IOException, InvalidFormatException {
            // If clearly doesn't do mark/reset, wrap up
            if (!inputStream.markSupported()) {
                inputStream = new PushbackInputStream(inputStream, 8);
            }

            // Ensure that there is at least some data there
            byte[] header8 = IOUtils.peekFirst8Bytes(inputStream);

            // Determine file type
            if (NPOIFSFileSystem.hasPOIFSHeader(header8)) {
                return XLS;
            } else {    // POIXMLDocument.hasOOXMLHeader(inputStream)
                return XLSX;
            }
        }

        /**
         * 返回Excel工作簿对象。
         *
         * @param inputStream Excel文件输入流
         * @return Excel工作簿对象
         * @throws IOException IO流不可读
         */
        public static org.apache.poi.ss.usermodel.Workbook getWorkbook(InputStream inputStream) throws IOException, InvalidFormatException {
            return WorkbookFactory.create(inputStream);
        }
    }
}
