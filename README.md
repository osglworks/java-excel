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

Excel reader is intelligent enough to skip non-matching worksheets, e.g. cover page. However it also provide ways for developer to manually select or exclude specific worksheets:

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

At the moment Excel reader does not support reading into nested POJO.  

## Caption to property mapping

By default caption will be mapped into standard Java property name, i.e. the lowercase started camel case string, e.g. `First name` will be mapped into `firstName`. This feature makes it very handy when reading into a POJO data list. However when reading into a Map list and you want to keep the caption as map key, then you can create `ExcelReader` using different caption to property transform strategy:

```java
List<Map<String, Object>> data = ExcelReader.builder(CaptionSchemaTransformStrategy.AS_CAPTION)
        .file(sampleFile())
        .build()
        .read();
```

In case your caption is in a different language, you must do manual map:

```java
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
```

Or you can do it this way:

```java
ExcelReader reader = ExcelReader.builder(
        C.map("姓", "lastName", "名", "firstName", ...)
).file(sampleFile()).build();
```

This technique also applied when your caption is English words but there is no simple way to process the transform through any strategy, e.g. `Street #` into `streetNo` etc.

## Type 