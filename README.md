# osgl-excel-reader

A flexible and easy to use java excel reading kit.

## Install

Add the following section to your `pom.xml` file:

```xml
<dependency>
  <groupId>org.osgl</groupId>
  <artifactId>excel-reader</artifactId>
  <version>1.0.0-SNAPSHOT</version>
</dependency>
```

**Note** If you are using SNAPSHOT version make sure you have the parent section in your `pom.xml` file:

```xml
<parent>
<groupId>org.sonatype.oss</groupId>
<artifactId>oss-parent</artifactId>
<version>7</version>
</parent>
```

## Simple use cases

Read all sheets into list of ordered map data structure:

```java
List<Map<String, Object>> data = ExcelReader.read(new File("/path/to/my/excel.xls"));
```

Read all sheets into list of POJO object:

```java
List<MyBean> data = ExcelReader.read(new File("/path/to/my/excel.xls"), MyBean.class);
```

## Include/exclude sheets

Excel reader is intelligent enough to skip non-matching worksheets, e.g. cover page. However it also provides ways for developer to manually select or exclude specific worksheets:

### Read specified sheets

```java
List<MyBean> data = ExcelReader.builder()
    .sheets("2016,2017", "2018", "summary", ...)
    .file(new File("/path/..."))
    .build().read(MyBean.class);
```

### exclude worksheets

```java
List<MyBean> data = ExcelReader.builder()
    .excludeSheets("Cover page, Summary")
    .file(new File("/path/..."))
    .build().read(MyBean.class);
```

## Type conversion

When reading into Map, Excel reader use the cell type to determine the read value type. Excel reader is smart enough to read Date value even it is typed as `Numeric` in excel.

When reading into POJO, Excel reader will cover the read value into the field type, including all primitive types, wrapped types, String and any enum type.

## Nested POJO

Excel reader support loading data into nested POJO. For example, suppose we have a nested POJO as:

```java
// The address 
public class Address {
    private String unitNo;
    ...
    public String getUnitNo() {
        return unitNo;
    }
    public void setUnitNo(String unitNo) {
        this.unitNo = unitNo;
    }
    ...
}
// The student
public class Student {
    public String id;
    public String name;
    public Address address;
}
```

And we have an excel sheet looks like

| Id  | Name | Unit # | Street # | Street | Suburb | State |
| --- | ---- | ------ | -------- | ------ | ----- | ----- |
| A001 | John Smith | 2 | 2-4 | Peak | CastleHill | NSW |

We can load the data into the student POJO as:

```java
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
``` 

**Note** we just need to map the headers that cannot be transformed automatically, for headers like `Id` and `Name` we don't need to map them because Excel reader can handle them automatically.

## Header mapping

By default header will be mapped into standard Java property name, i.e. the lowercase started camel case string, e.g. `First name` will be mapped into `firstName`. This feature makes it very handy when reading into a POJO data list. However when reading into a Map list and you want to keep the header as map key, then you can create `ExcelReader` using different header to property transform strategy:

```java
List<Map<String, Object>> data = ExcelReader.builder(HeaderSchemaTransformStrategy.AS_CAPTION)
        .file(sampleFile())
        .build()
        .read();
```

In case your header is in a different language, you must do manual mapping:

```java
ExcelReader reader = ExcelReader.builder()
        .map("姓").to("lastName")
        .map("名").to("firstName")
        .map("学号").to("no")
        .map("出生日期").to("dob")
        .map("年级").to("grade")
        .map("国家").to("country")
        .map("邮编").to("postCode")
        .file(sampleFile())
        .build();
List<Student> data = reader.read(Student.class);
```

Or you can do it this way:

```java
ExcelReader reader = ExcelReader.builder(
        C.map("姓", "lastName", "名", "firstName", ...)
).file(sampleFile()).build();
```

This technique also applied when your header is English words but there is no simple way to process the transform through any strategy, e.g. `Street #` into `streetNo` etc.

