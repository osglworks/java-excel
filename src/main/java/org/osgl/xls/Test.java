package org.osgl.xls;

import org.osgl.util.C;

import java.io.File;
import java.util.List;
import java.util.Map;

public class Test {

    public static class Foo {
        public String qtId;
        public String subject;
        public String module;
        public String assignee;

        @Override
        public String toString() {
            return "Foo{" +
                    "qtId='" + qtId + '\'' +
                    ", subject='" + subject + '\'' +
                    ", module='" + module + '\'' +
                    ", assignee='" + assignee + '\'' +
                    '}';
        }
    }

    public static void main(String[] args) {
        File file = new File("/home/luog/Downloads/test.xls");
        List<Foo> dataList1 = ExcelReader.builder(C.map("Assigned To", "assignee",
                "Summary", "subject",
                "Id", "qtId"
        ))
                .file(file)
                .build().read(Foo.class);
        for (Foo foo : dataList1) {
            System.out.println(foo);
        }
        List<Map<String, Object>> dataList2 = ExcelReader.builder().file(file).build().read();
        for (Map<String, Object> map : dataList2) {
            System.out.println(map);
        }
    }
}
