package com.convert;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.usermodel.Cell.*;

/**
 * Excel文件读取
 *
 * Created by Dengbin on 2015/10/4.
 */
public class DefaultExcelReader implements ExcelReader {
    private static final Logger log = LogManager.getLogger();

    @Override
    public String[] readHeader(InputStream inputStream) {
        return new String[0];
    }

    @Override
    public String[] readHeader(Workbook workbook) {
        return new String[0];
    }

    @Override
    public <T> T[] readContent(InputStream inputStream, Class<T> clazz) throws IOException, InvalidFormatException, InstantiationException, IllegalAccessException, InvocationTargetException {
        return readContent(WorkbookFactory.create(inputStream), clazz);
    }

    @Override
    public <T> T[] readContent(Workbook workbook, Class<T> clazz) throws IllegalAccessException, InstantiationException, InvocationTargetException, IOException, InvalidFormatException {
        AnnotationParser parser = new ExcelAnnotationParser();
        Map<Integer, Map<String, Object>> contents = getSheetContent(workbook);
        Map<String, String> dictionary = parser.parse(Clerk.class);

        return mapper(clazz, contents, dictionary);
    }

    private Map<Integer, Map<String, Object>> getSheetContent(Workbook workbook) throws IOException, InvalidFormatException, IllegalAccessException, InvocationTargetException, InstantiationException {
        return mergeContent(readContent(workbook));
    }

    /**
     * 将Excel数据匹配到Java对象中。
     *
     * @param clazz Java对象类型
     * @param contents 需要进行匹配的数据
     * @param dictionary 进行匹配时的字段对照字典
     * @param <T> Java对象类型，用来强制转换数组类型
     * @return Java对象数组
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    @SuppressWarnings("unchecked")
    public <T> T[] mapper(Class<T> clazz, Map<Integer, Map<String, Object>> contents, Map<String, String> dictionary) throws IllegalAccessException, InstantiationException {
        Map<String, Field> fields = new HashMap<>();
        for (Field field : clazz.getDeclaredFields()) {
            fields.put(field.getName(), field);
        }
        List<T> list = new ArrayList<>();
        for (Map<String, Object> content : contents.values()) {
            T obj = clazz.newInstance();
            mapper(obj, fields, content, dictionary);
            list.add(obj);
        }
        // type safety Object[]
        T[] ts = (T[]) Array.newInstance(clazz, list.size());
        for (int i = 0; i < ts.length; i++) {
            ts[i] = list.get(i);
        }
        return ts;
    }

    /**
     * 将Excel数据匹配到Java对象中。
     *
     * @param obj 需要被赋值的对象
     * @param fields 需要被赋值的字段
     * @param content 为字段提供值的MAP集合
     * @param dictionary 字段名与Excel列名的对照表
     * @param <T> 需要被赋值的对象的类型
     * @throws IllegalAccessException
     */
    private <T> void mapper(T obj, Map<String, Field> fields, Map<String, Object> content, Map<String, String> dictionary) throws IllegalAccessException {
        for (Map.Entry<String, Field> fieldEntry : fields.entrySet()) {
            fieldEntry.getValue().setAccessible(true);
            fieldEntry.getValue().set(obj, content.get(dictionary.get(fieldEntry.getKey())));
        }
    }

    /**
     * 当Excel工作簿中的表格都为同一类型的数据时将其合并方便操作。
     *
     * @param contents Excel工作簿的所有数据集合
     * @return Excel工作簿的数据集合，但不区分表格名称。
     */
    private Map<Integer, Map<String, Object>> mergeContent(Map<String, Map<Integer, Map<String, Object>>> contents) {
        Map<Integer, Map<String, Object>> all = new LinkedHashMap<>();
        int index = 1;
        for (Map<Integer, Map<String, Object>> map : contents.values()) {
            for (Map<String, Object> content : map.values()) {
                all.put(index++, content);
            }
        }

        return all;
    }

    /**
     * 返回整个工作簿的所有数据集合。
     *
     * @param workbook Excel工作簿
     * @return Excel工作簿的数据集合，并区分表格名称。
     */
    private Map<String, Map<Integer, Map<String, Object>>> readContent(Workbook workbook) {
        int totalContents = 0;
        int sheetsNum = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetsNum; i++) {
            totalContents += workbook.getSheetAt(i).getLastRowNum();
        }
        Map<String, Map<Integer, Map<String, Object>>> allContents = new LinkedHashMap<>(totalContents);

        for (int i = 0; i < sheetsNum; i++) {
            allContents.put(workbook.getSheetAt(i).getSheetName(), readContent(workbook.getSheetAt(i)));
        }

        return allContents;
    }

    /**
     * 返回Excel中某一张表格的所有数据集合。
     *
     * @param sheet Excel表格
     * @return Excel表格的数据集合
     */
    private Map<Integer, Map<String, Object>> readContent(Sheet sheet) {
        Map<Integer, Map<String, Object>> contents = new LinkedHashMap<>(sheet.getLastRowNum());
        Row headerRow = sheet.getRow(0);    // 表头
        // Excel表的总列数
        int colNum = headerRow.getPhysicalNumberOfCells();

        // 获取总行数
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row contentRow = sheet.getRow(i);
            Map<String, Object> rowContents = new LinkedHashMap<>(1);
            for (int col = 0; col < colNum; col++) {
                rowContents.put(headerRow.getCell(col).toString(), getCellValue(contentRow.getCell(col)));
            }
            contents.put(i, rowContents);
        }

        return contents;
    }

    /**
     * 返回Excel单元格的数据。
     *
     * @param cell Excel单元格
     * @return 单元格的数据
     */
    private Object getCellValue(Cell cell) {
        return getCellValue(cell, "yyyy-MM-dd");
    }

    /**
     * 返回Excel单元格的数据。
     *
     * @param cell Excel单元格
     * @return 单元格的数据
     */
    private Object getCellValue(Cell cell, String format) {
        return getCellValue(cell, new SimpleDateFormat(format));
    }

    /**
     * 返回Excel单元格的数据。
     *
     * @param cell Excel单元格
     * @return 单元格的数据
     */
    private <T extends DateFormat> String getCellValue(Cell cell, T formatter) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case CELL_TYPE_BLANK:
                return "";
            case CELL_TYPE_BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {   // 当公式值为Date时
                    return formatter.format(cell.getDateCellValue());
                }
                return Double.toString(cell.getNumericCellValue());
            case CELL_TYPE_STRING:
                return cell.getStringCellValue().trim();
            case CELL_TYPE_ERROR:
                log.warn("数据读取有误[{}]！", cell.toString());
                return null;
            case CELL_TYPE_FORMULA:
                try {
                    if (DateUtil.isCellDateFormatted(cell)) {   // 当公式值为Date时
                        return formatter.format(cell.getDateCellValue());
                    }
                    return Double.toString(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    try {
                        return cell.getStringCellValue().trim();
                    } catch (Exception ex) {
                        return null;
                    }
                }
            default: return null;   // not happened
        }
    }
}
