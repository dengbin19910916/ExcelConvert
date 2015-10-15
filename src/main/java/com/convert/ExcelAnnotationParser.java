package com.convert;

import com.convert.annotation.Title;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Dengbin on 2015/10/6.
 */
public class ExcelAnnotationParser implements AnnotationParser {

    /**
     * 解析类型中所有的注解信息。
     *
     * @param clazz 需要解析注解信息的类型
     */
    @SuppressWarnings("unchecked")
    @Override
    public Map<String, String> parse(Class clazz) {
        return getTitleDictionary(clazz);
    }

    /**
     * 返回属性名称所对应的中文名称，即Excel的列名。
     *
     * @param clazz Excel工作簿的映射对象类型
     * @return 属性名称与Excel列名的对应关系
     */
    public Map<String, String> getTitleDictionary(Class clazz) {
        Map<String, String> dictionary = new HashMap<>();
        Method[] methods = Arrays.asList(clazz.getDeclaredMethods())
                .stream()
                .filter(method -> method.getName().startsWith("set"))
                .toArray(Method[]::new);
        for (Method method : methods) {
            Title title = method.getDeclaredAnnotation(Title.class);
            if (title != null) {
                dictionary.put(method.getName().substring(3, 4).toLowerCase() + method.getName().substring(4), title.name());
            }
        }

        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            Title title = field.getDeclaredAnnotation(Title.class);
            if (title != null) {
                dictionary.merge(field.getName(), title.name(), (oldVal, newVal) -> newVal);
            }
        }

        return dictionary;
    }

    private static class Workbook0 {
        private String name;
        private Sheet0[] sheets;
        private String type;

        public Workbook0() {
            super();
        }

        public Workbook0(String name, Sheet0[] sheets, String type) {
            this.name = name;
            this.sheets = sheets;
            this.type = type;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Sheet0[] getSheets() {
            return sheets;
        }

        public void setSheets(Sheet0[] sheets) {
            this.sheets = sheets;
        }

        public String getType() {
            return type;
        }

        public void setType(String type) {
            this.type = type;
        }

        @Override
        public String toString() {
            return "Workbook0{" +
                    "name='" + name + '\'' +
                    ", sheets=" + Arrays.toString(sheets) +
                    ", type='" + type + '\'' +
                    '}';
        }
    }

    private static class Sheet0 {
        private String name;
        private List<Title> titles;

        public Sheet0() {
            super();
        }

        public Sheet0(String name, List<Title> titles) {
            this.name = name;
            this.titles = titles;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public List<Title> getTitles() {
            return titles;
        }

        public void setTitles(List<Title> titles) {
            this.titles = titles;
        }

        @Override
        public String toString() {
            return "Sheet0{" +
                    "name='" + name + '\'' +
                    ", titles=" + titles +
                    '}';
        }
    }

    private static class Title0 {
        private String name;

        public Title0() {
            super();
        }

        public Title0(String name) {
            this.name = name;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        @Override
        public String toString() {
            return "Title0{" +
                    "name='" + name + '\'' +
                    '}';
        }
    }
}
