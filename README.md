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

## Sample usage

Read all sheets into list of ordered map data structure:

```java
List<Map<String, Object>> data = ExcelReader.read(new File("/path/to/my/excel.xls"));
```

Read all sheets into list of POJO object:

```java
List<MyBean> data = ExcelReader.read(new File("/path/to/my/excel.xls"), MyBean.class);
```

Read sheets with caption to field/map key transform

```java
List<MyBean> data = ExcelReader.read(
        new File("/path/to/my/excel.xls"), 
        MyBean.class,
        C.map("姓名", "name", "性别", "gender", ...));
```

Read specified sheets

```java
List<MyBean> data = ExcelReader.builder()
    .sheets("2016,2017", "2018", "summary", ...)
    .file(new File("/path/..."))
    .build().read(MyBean.class);
```

Read with sheets excluded

```java
List<MyBean> data = ExcelReader.builder()
    .excludeSheets("Cover page, Summary")
    .file(new File("/path/..."))
    .build().read(MyBean.class);
```
