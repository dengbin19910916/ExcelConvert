package com.convert;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.List;

import static com.convert.annotation.Workbook.Type;

public class Test {

    public static void main(String[] args) throws IOException, InvocationTargetException, InvalidFormatException, InstantiationException, IllegalAccessException {
        ExcelReader reader = new DefaultExcelReader();
        InputStream inputStream = new FileInputStream("C:\\WorkSpace\\test2.xlsx");
        Clerk[] clerks = reader.readContent(inputStream, Clerk.class);

        int i = 1;
        for (Object o : clerks) {
            System.out.println(i++ + "----" + o);
        }

        /*Class c1 = int.class;
        Class c2 = Integer.class;
        Class c3 = int.class;
        Class c4 = Integer.class;

        System.out.println(c1 == c2);
        System.out.println(c1 == c3);
        System.out.println(c2 == c4);*/

        /*Gson gson = new Gson();
        String src = "{\"name\":\"行员管理\",\"sheets\":[{\"name\":\"城东支行\",\"titles\":[{\"name\":\"行员代号\"},{\"name\":\"行员名称\"},{\"name\":\"出生日期\"}]},{\"name\":\"城西支行\",\"titles\":[{\"name\":\"行员代号\"},{\"name\":\"行员名称\"},{\"name\":\"出生日期\"}]}],\"type\":\"XLSX\"}";
        List<Title0> titles = new ArrayList<Title0>() {{
            add(new Title0("行员代号"));
            add(new Title0("行员名称"));
            add(new Title0("出生日期"));
        }};
        Sheet0[] sheets = new Sheet0[] {new Sheet0("城东支行", titles), new Sheet0("城西支行", titles)};

        Workbook0 workbook = new Workbook0("行员管理", Type.XLSX, sheets);
        String json = gson.toJson(workbook);
        System.out.println(json);

        Workbook0 wb = gson.<Workbook0>fromJson(src, new TypeToken<Workbook0>(){}.getType());
        Workbook0 wb2 = gson.fromJson(src, Workbook0.class);
        System.out.println(wb);*/
    }

    private static class Workbook0 {

        /**
         * Excel 工作簿的名称，用于生成Excel文件。
         */
        private String name;
        /**
         * Excel 工作簿的保存类型，默认为XLS，用于生成Excel文件。
         */
        private Type type;
        /**
         * Excel 工作簿中的表格组。
         */
        private Sheet0[] sheets;

        public Workbook0() {
            super();
        }

        public Workbook0(String name, Type type, Sheet0[] sheets) {
            this.name = name;
            this.type = type;
            this.sheets = sheets;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Type getType() {
            return type;
        }

        public void setType(Type type) {
            this.type = type;
        }

        public Sheet0[] getSheets() {
            return sheets;
        }

        public void setSheets(Sheet0[] sheets) {
            this.sheets = sheets;
        }

        @Override
        public String toString() {
            return "Workbook0{" +
                    "name='" + name + '\'' +
                    ", type=" + type +
                    ", sheets=" + Arrays.toString(sheets) +
                    '}';
        }
    }

    private static class Sheet0 {
        /**
         * 表格名称。
         */
        private String name;
        /**
         * 表格中的表头。
         */
        private List<Title0> titles;

        public Sheet0() {
            super();
        }

        public Sheet0(String name, List<Title0> titles) {
            this.name = name;
            this.titles = titles;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public List<Title0> getTitles() {
            return titles;
        }

        public void setTitles(List<Title0> titles) {
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
        /**
         * 表头名称
         */
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
