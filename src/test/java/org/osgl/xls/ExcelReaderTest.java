package org.osgl.xls;

import org.junit.Assert;
import org.junit.Test;
import org.osgl.util.E;

import java.io.File;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class ExcelReaderTest extends Assert {

    private URL sampleUrl() {
        return ExcelReaderTest.class.getResource("/students.xlsx");
    }

    private File sampleFile() {
        return new File(sampleUrl().getFile());
    }

    @Test
    public void testReadIntoMap() {
        List<Map<String, Object>> data = ExcelReader.read(sampleFile());
        assertEquals(2, data.size());
        Map<String, Object> s1 = data.get(0);
        assertEquals("张", s1.get("姓"));
        assertEquals("武侯区", s1.get("区"));
        Map<String, Object> s2 = data.get(1);
        assertEquals("John", s2.get("firstName"));
        assertEquals("Premier St", s2.get("street"));
    }

    @Test
    public void testReadIntoMapWithKeepCaptionSchemaTransformStategy() {
        List<Map<String, Object>> data = ExcelReader.builder(HeaderTransformStrategy.AS_CAPTION)
                .file(sampleFile())
                .build()
                .read();
        Map<String, Object> s1 = data.get(0);
        assertEquals("张", s1.get("姓"));
        assertEquals("武侯区", s1.get("区"));
        assertEquals(dobOfZhang(), s1.get("出生日期"));
        Map<String, Object> s2 = data.get(1);
        assertEquals("John", s2.get("First Name"));
        assertEquals("Premier St", s2.get("Street"));
        assertEquals(dobOfJohn(), s2.get("DOB"));
    }

    @Test
    public void testReadIntoPojo() {
        List<Student> data = ExcelReader.read(sampleFile(), Student.class);
        assertEquals(1, data.size());
        Student s1 = data.get(0);
        assertEquals("John", s1.getFirstName());
        assertEquals(Student.Grade.g1, s1.getGrade());
        assertEquals(Student.Country.AUSTRALIA, s1.getCountry());
        assertEquals(dobOfJohn(), s1.getDob());
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
        assertEquals(2, data.size());

        Student s0 = data.get(0);
        assertEquals("三", s0.getFirstName());
        assertEquals(Student.Grade.g1, s0.getGrade());
        assertEquals(Student.Country.CHINA, s0.getCountry());
        assertEquals(dobOfZhang(), s0.getDob());

        Student s1 = data.get(1);
        assertEquals("John", s1.getFirstName());
        assertEquals(Student.Grade.g1, s1.getGrade());
        assertEquals(Student.Country.AUSTRALIA, s1.getCountry());
        assertEquals(dobOfJohn(), s1.getDob());
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
        assertEquals(1, data.size());
        Student s1 = data.get(0);
        Address addr = s1.getAddress();
        assertNotNull(addr);
        assertEquals("29", addr.getUnitNo());
        assertEquals("26-30", addr.getStreetNo());
        assertEquals("Premier St", addr.getStreet());
        assertEquals("Kogarah", addr.getSuburb());
        assertEquals("NSW", addr.getState());
        assertEquals("2217", addr.getPostCode());
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
