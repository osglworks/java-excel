package org.osgl.xls;

/*-
 * #%L
 * Java Excel Reader
 * %%
 * Copyright (C) 2017 OSGL (Open Source General Library)
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

import org.junit.Test;
import org.osgl.ut.TestBase;
import org.osgl.util.E;

import java.io.File;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class ExcelReaderTest extends TestBase {

    private URL sampleUrl() {
        return ExcelReaderTest.class.getResource("/students.xlsx");
    }

    private File sampleFile() {
        return new File(sampleUrl().getFile());
    }

    @Test
    public void testReadIntoMap() {
        List<Map<String, Object>> data = ExcelReader.read(sampleFile());
        eq(2, data.size());
        Map<String, Object> s1 = data.get(0);
        eq("张", s1.get("姓"));
        eq("武侯区", s1.get("区"));
        Map<String, Object> s2 = data.get(1);
        eq("John", s2.get("firstName"));
        eq("Premier St", s2.get("street"));
    }

    @Test
    public void testReadIntoMapWithKeepCaptionSchemaTransformStategy() {
        List<Map<String, Object>> data = ExcelReader.builder(HeaderTransformStrategy.AS_CAPTION)
                .file(sampleFile())
                .build()
                .read();
        Map<String, Object> s1 = data.get(0);
        eq("张", s1.get("姓"));
        eq("武侯区", s1.get("区"));
        eq(dobOfZhang(), s1.get("出生日期"));
        Map<String, Object> s2 = data.get(1);
        eq("John", s2.get("First Name"));
        eq("Premier St", s2.get("Street"));
        eq(dobOfJohn(), s2.get("DOB"));
    }

    @Test
    public void testReadIntoPojo() {
        List<Student> data = ExcelReader.read(sampleFile(), Student.class);
        eq(1, data.size());
        Student s1 = data.get(0);
        eq("John", s1.getFirstName());
        eq(Student.Grade.g1, s1.getGrade());
        eq(Student.Country.AUSTRALIA, s1.getCountry());
        eq(dobOfJohn(), s1.getDob());
    }

    @Test
    public void testReadIntoPojoWithSchemaTransform() {
        ExcelReader reader = ExcelReader.builder()
                .map("姓").to("lastName")
                .map("名").to("firstName")
                .map("ID").to("no")
                .map("学号").to("no")
                .map("出生日期").to("dob")
                .map("年级").to("grade")
                .map("国家").to("country")
                .map("邮编").to("postCode")
                .file(sampleFile())
                .build();
        List<Student> data = reader.read(Student.class);
        eq(2, data.size());

        Student s0 = data.get(0);
        eq("三", s0.getFirstName());
        eq(Student.Grade.g1, s0.getGrade());
        eq(Student.Country.CHINA, s0.getCountry());
        eq(dobOfZhang(), s0.getDob());

        Student s1 = data.get(1);
        eq("John", s1.getFirstName());
        eq(Student.Grade.g1, s1.getGrade());
        eq(Student.Country.AUSTRALIA, s1.getCountry());
        eq(dobOfJohn(), s1.getDob());
    }

    @Test
    public void testReadIntoNestedPojo() {
        ExcelReader reader = ExcelReader.builder()
                .map("Unit #").to("address.unitNo")
                .map("Street #").to("address.streetNo")
                .map("Street").to("address.street")
                .map("Suburb").to("address.suburb")
                .map("State").to("address.state")
                .map("Post code").to("address.postCode")
                .file(sampleFile())
                .build();
        List<Student> data = reader.read(Student.class);
        eq(1, data.size());
        Student s1 = data.get(0);
        Address addr = s1.getAddress();
        notNull(addr);
        eq("29", addr.getUnitNo());
        eq("26-30", addr.getStreetNo());
        eq("Premier St", addr.getStreet());
        eq("Kogarah", addr.getSuburb());
        eq("NSW", addr.getState());
        eq("2217", addr.getPostCode());
    }

    private static DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

    private static Date dobOfZhang() {
        return dateOf("2006-05-23");
    }

    private static Date dobOfJohn() {
        return dateOf("2006-07-20");
    }

    private static Date dateOf(String dateStr) {
        try {
            return dateFormat.parse(dateStr);
        } catch (Exception e) {
            throw E.unexpected(e);
        }
    }

}
