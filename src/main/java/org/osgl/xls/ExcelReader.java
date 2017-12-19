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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.osgl.$;
import org.osgl.Osgl;
import org.osgl.exception.NotAppliedException;
import org.osgl.logging.LogManager;
import org.osgl.logging.Logger;
import org.osgl.storage.ISObject;
import org.osgl.util.*;
import osgl.version.Version;

import java.io.*;
import java.util.*;

public class ExcelReader {

    public static final Version VERSION = Version.of(ExcelReader.class);

    public static final Logger LOGGER = LogManager.get(ExcelReader.class);

    private final boolean isXlsx;
    private final $.Func0<InputStream> inputStreamProvider;
    private final $.Predicate<Sheet> sheetSelector;
    private final int headerRow;
    private final boolean ignoreEmptyRows;
    private final TolerantLevel tolerantLevel;
    private final $.Function<String, String> headerTransformer;
    private final Map<String, String> headerMapping;
    private final String terminator;

    private ExcelReader(Builder builder) {
        inputStreamProvider = $.notNull(builder.inputStreamProvider);
        sheetSelector = builder.sheetSelector;
        headerRow = builder.headerRow;
        ignoreEmptyRows = builder.ignoreEmptyRows;
        isXlsx = builder.isXlsx;
        tolerantLevel = builder.tolerantLevel;
        headerMapping = builder.headerMapping;
        headerTransformer = builder.headerTransformer;
        terminator = builder.terminator;
    }

    public List<Map<String, Object>> read() {
        return (List) read(Map.class);
    }

    public <TYPE> List<TYPE> read(Class<? extends TYPE> schema) {
        final List<TYPE> dataList = new ArrayList<>();
        long now = $.ms();
        final Workbook wb = loadWorkbook();
        long time = $.ms() - now;
        LOGGER.trace("it takes %sms to load the workbook", time);
        try {
            Map<String, PropertySetter> setterMap = processSchemaMapping(schema);
            if (setterMap.isEmpty() && tolerantLevel.isStrict()) {
                throw new ExcelReadException("No schema mapping found in strict mode");
            }
            read(wb, dataList, setterMap, schema);
        } finally {
            IO.close(wb);
        }
        return dataList;
    }

    private <TYPE> Map<String, PropertySetter> processSchemaMapping(Class<? extends TYPE> schema) {
        Map<String, PropertySetter> schemaMapping = new HashMap<>();
        final boolean schemaIsPojo = !Map.class.isAssignableFrom(schema);
        if (schemaIsPojo) {
            Map<String, PropertySetter> pojoSetters = JavaBeanUtil.setters(schema, headerMapping);
            schemaMapping.putAll(pojoSetters);
            for (Map.Entry<String, String> entry : headerMapping.entrySet()) {
                String header = entry.getKey().trim().toLowerCase();
                String field = entry.getValue();
                PropertySetter setter = pojoSetters.get(field);
                if (null != setter) {
                    schemaMapping.put(header, setter);
                } else if (tolerantLevel.isTolerant()) {
                    field = headerTransformer.apply(header);
                    setter = pojoSetters.get(field);
                    if (null != setter) {
                        schemaMapping.put(header, setter);
                    }
                }
            }
        } else {
            for (Map.Entry<String, String> entry : headerMapping.entrySet()) {
                schemaMapping.put(entry.getKey(), new MapSetter(entry.getValue()));
            }
        }
        return schemaMapping;
    }

    private <TYPE> void read(Workbook workbook, final List<TYPE> dataList, Map<String, PropertySetter> setterMap, Class<? extends TYPE> schema) {
        for (Sheet sheet : workbook) {
            if (!sheetSelector.test(sheet)) {
                continue;
            }
            read(sheet, dataList, setterMap, schema);
        }
    }

    private <TYPE> void read(Sheet sheet, final List<TYPE> dataList, Map<String, PropertySetter> setterMap, Class<? extends TYPE> schema) {
        $.Var<Integer> headerRowHolder = $.var(this.headerRow);
        boolean schemaIsMap = Map.class.isAssignableFrom(schema);
        Map<Integer, PropertySetter> columnIndex = buildColumnIndex(sheet, setterMap, schemaIsMap, headerRowHolder);
        if (columnIndex.size() < setterMap.size()) {
            tolerantLevel.columnIndexMapNotFullyBuilt(sheet);
        }
        if (columnIndex.isEmpty()) {
            return;
        }
        int startRow = headerRowHolder.get() + 1;
        int maxRow = sheet.getLastRowNum() + 1;
        TERMINATED:
        for (int rowId = startRow; rowId < maxRow; ++rowId) {
            Object entity = schemaIsMap ? new LinkedHashMap<>() : $.newInstance(schema);
            Row row = sheet.getRow(rowId);
            boolean isEmptyRow = true;
            for (Map.Entry<Integer, PropertySetter> entry : columnIndex.entrySet()) {
                try {
                    Cell cell = row.getCell(entry.getKey());
                    try {
                        Object value = readCellValue(cell);
                        if (null != value) {
                            if (null != terminator && terminator.equals(value)) {
                                break TERMINATED;
                            }
                            isEmptyRow = false;
                            try {
                                entry.getValue().set(entity, value, null);
                            } catch (Exception e) {
                                tolerantLevel.errorSettingCellValueToPojo(e, cell, value, schema);
                            }
                        }
                    } catch (Exception e) {
                        tolerantLevel.onReadCellException(e, cell);
                    }
                } catch (Exception e) {
                    tolerantLevel.onReadCellException(e, sheet, row, entry.getKey());
                }
            }
            if (isEmptyRow && ignoreEmptyRows) {
                continue;
            }
            TYPE data = $.cast(entity);
            dataList.add(data);
        }
    }

    private Object readCellValue(Cell cell) {
        return readCellValue(cell, cell.getCellTypeEnum());
    }

    private Object readCellValue(Cell cell, CellType type) {
        switch (type) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                double n = cell.getNumericCellValue();
                CellStyle style = cell.getCellStyle();
                String format = style.getDataFormatString();
                if (format.indexOf('.') < 0) {
                    return (long) n;
                }
                return n;
            case FORMULA:
                return readCellValue(cell, cell.getCachedFormulaResultTypeEnum());
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                return tolerantLevel.readErrorCell(cell);
            case BLANK:
                return null;
            case STRING:
                return cell.getStringCellValue();
            default:
                return tolerantLevel.readUnknownCellType(cell);
        }
    }

    private Map<Integer, PropertySetter> buildColumnIndex(Sheet sheet, Map<String, PropertySetter> setterMap, boolean schemaIsMap, $.Var<Integer> headerRowHolder) {
        int startRow = sheet.getFirstRowNum();
        int maxRow = sheet.getLastRowNum();
        int headerRow = this.headerRow;
        if (headerRow < startRow || headerRow >= maxRow) {
            tolerantLevel.headerRowOutOfScope(sheet);
            headerRow = startRow;
        }
        if (tolerantLevel.isStrict()) {
            return buildColumnIndex(sheet.getRow(headerRow), setterMap, schemaIsMap);
        }
        Map<Integer, PropertySetter> index = C.map();
        for (int rowId = headerRow; rowId < maxRow; ++rowId) {
            index = buildColumnIndex(sheet.getRow(rowId), setterMap, schemaIsMap);
            if (!index.isEmpty()) {
                headerRowHolder.set(rowId);
                return index;
            }
        }
        return index;
    }

    private Map<Integer, PropertySetter> buildColumnIndex(Row row, Map<String, PropertySetter> setterMap, boolean schemaIsMap) {
        Map<Integer, PropertySetter> retVal = new HashMap<>();
        for (Cell cell : row) {
            try {
                String header = cell.getStringCellValue();
                if (S.blank(header)) {
                    continue;
                }
                header = header.trim();
                PropertySetter setter = null;
                String translated = headerMapping.get(header.toLowerCase());
                if (null != translated) {
                    setter = setterMap.get(translated);
                }
                if (null == setter) {
                    String key = headerTransformer.apply(header);
                    setter = setterMap.get(key);
                }
                if (null != setter) {
                    retVal.put(cell.getColumnIndex(), setter);
                } else if (tolerantLevel.isAggressiveReading() && schemaIsMap) {
                    retVal.put(cell.getColumnIndex(), new MapSetter(headerTransformer.apply(header)));
                }
            } catch (Exception e) {
                LOGGER.debug(e, "error reading cell value:" + cell);
            }
        }
        return retVal;
    }

    private Workbook loadWorkbook() {
        InputStream is = inputStreamProvider.apply();
        try {
            return isXlsx ? new XSSFWorkbook(is) : new HSSFWorkbook(is);
        } catch (IOException e) {
            throw E.ioException(e);
        }
    }

    public enum TolerantLevel {
        STRICT, TOLERANT, AGGRESSIVE_READ;

        public boolean isStrict() {
            return STRICT == this;
        }

        public boolean isTolerant() {
            return TOLERANT != this;
        }

        public boolean isAggressiveReading() {
            return AGGRESSIVE_READ == this;
        }

        public void headerRowOutOfScope(Sheet sheet) {
            String message = S.fmt("caption row out of scope in sheet[%s] !", sheet.getSheetName());
            if (isStrict()) {
                throw new ExcelReadException(message);
            }
            LOGGER.warn(message + " will probe for caption row");
        }

        public void columnIndexMapNotFullyBuilt(Sheet sheet) {
            String message = S.fmt("column index not fully built on sheet: " + sheet.getSheetName());
            if (isStrict()) {
                throw new ExcelReadException(message);
            }
            LOGGER.warn(message);
        }

        public void onReadCellException(Exception e, Sheet sheet, Row row, int cellIndex) {
            String errorMessage = S.fmt("Error reading cell value: %s-%s@[%s]", cellIndex, row.getRowNum(), sheet.getSheetName());
            if (isStrict()) {
                throw new ExcelReadException(e, errorMessage);
            }
            LOGGER.warn(e, errorMessage);
        }

        public void onReadCellException(Exception e, Cell cell) {
            String errorMessage = S.fmt("Error reading cell value: %s@[%s]", cell.getAddress(), cell.getSheet().getSheetName());
            if (isStrict()) {
                throw new ExcelReadException(e, errorMessage);
            }
            LOGGER.warn(e, errorMessage);
        }

        public Object readErrorCell(Cell cell) {
            if (isStrict()) {
                throw new ExcelReadException("Error cell value encountered: %s@[%s]", cell.getAddress(), cell.getRow().getSheet().getSheetName());
            }
            return null;
        }

        public Object readUnknownCellType(Cell cell) {
            if (isStrict()) {
                throw new ExcelReadException("Unknown cell type encountered: %s@[%s]", cell.getAddress(), cell.getRow().getSheet().getSheetName());
            }
            return null;
        }

        public void errorSettingCellValueToPojo(Exception e, Cell cell, Object value, Class<?> schema) {
            String errorMessage = S.fmt("failed to set cell value[%s] to POJO[%s]: %s@[%s]", value, schema, cell.getAddress(), cell.getRow().getSheet().getSheetName());
            if (isStrict()) {
                throw new ExcelReadException(e, errorMessage);
            }
            LOGGER.warn(e, errorMessage);
        }
    }

    /**
     * A quick API for reading an excel file and return a {@link List list} of {@link Map maps}
     * where each map corresponding to an excel sheet row
     *
     * @param file the excel source file
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(File file) {
        return builder().file(file).build().read();
    }

    /**
     * A quick API for reading an excel file and return a {@link List list} of {@link Map maps}
     * where each map corresponding to an excel sheet row.
     *
     * A `headerMapping` {@link Map map} is provided to allow developer to translate
     * the header string into map key, e.g. `姓名` to `name`
     *
     * @param file          the excel source file
     * @param headerMapping a string-string map that defines caption to map key transform
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(File file, Map<String, String> headerMapping) {
        return builder().file(file).headerMapping(headerMapping).build().read();
    }

    /**
     * A quick API for reading an excel file and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * @param file   the excel source file
     * @param schema specify the POJO object type
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(File file, Class<T> schema) {
        return builder().file(file).build().read(schema);
    }

    /**
     * A quick API for reading an excel file and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * A `headerMapping` {@link Map map} is provided to allow developer to translate
     * the header string into map key, e.g. `姓名` to `name`
     *
     * @param file          the excel source file
     * @param schema        specify the POJO object type
     * @param headerMapping a string-string map that defines header to map key transform
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(File file, Class<T> schema, Map<String, String> headerMapping) {
        return builder().file(file).headerMapping(headerMapping).build().read(schema);
    }


    /**
     * A quick API for reading an excel inputStream and return a {@link List list} of {@link Map maps}
     * where each map corresponding to an excel sheet row
     *
     * @param inputStream the excel source inputStream
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(InputStream inputStream) {
        return builder().inputStream(inputStream).build().read();
    }

    /**
     * A quick API for reading an excel inputStream and return a {@link List list} of {@link Map maps}
     * where each map corresponding to an excel sheet row.
     *
     * A `headerMapping` {@link Map map} is provided to allow developer to translate
     * the header string into map key, e.g. `姓名` to `name`
     *
     * @param inputStream   the excel source inputStream
     * @param headerMapping a string-string map that defines header to map key transform
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(InputStream inputStream, Map<String, String> headerMapping) {
        return builder().inputStream(inputStream).headerMapping(headerMapping).build().read();
    }

    /**
     * A quick API for reading an excel inputStream and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * @param inputStream the excel source inputStream
     * @param schema      specify the POJO object type
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(InputStream inputStream, Class<T> schema) {
        return builder().inputStream(inputStream).build().read(schema);
    }

    /**
     * A quick API for reading an excel inputStream and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * A `headerMapping` {@link Map map} is provided to allow developer to translate
     * the header string into map key, e.g. `姓名` to `name`
     *
     * @param inputStream   the excel source inputStream
     * @param schema        specify the POJO object type
     * @param headerMapping a string-string map that defines caption to map key transform
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(InputStream inputStream, Class<T> schema, Map<String, String> headerMapping) {
        return builder().inputStream(inputStream).headerMapping(headerMapping).build().read(schema);
    }


    /**
     * A quick API for reading an excel sobj and return a {@link List list} of {@link Map maps}
     * where each map corresponding to an excel sheet row
     *
     * @param sobj the excel source sobj
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(ISObject sobj) {
        return builder().sobject(sobj).build().read();
    }

    /**
     * A quick API for reading an excel sobj and return a {@link List list} of {@link Map maps}
     * where each map corresponding to an excel sheet row.
     *
     * A `headerMapping` {@link Map map} is provided to allow developer to translate
     * the caption string into map key, e.g. `姓名` to `name`
     *
     * @param sobj          the excel source sobj
     * @param headerMapping a string-string map that defines caption to map key transform
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(ISObject sobj, Map<String, String> headerMapping) {
        return builder().sobject(sobj).headerMapping(headerMapping).build().read();
    }

    /**
     * A quick API for reading an excel sobj and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * @param sobj   the excel source sobj
     * @param schema specify the POJO object type
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(ISObject sobj, Class<T> schema) {
        return builder().sobject(sobj).build().read(schema);
    }

    /**
     * A quick API for reading an excel sobj and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * A `headerMapping` {@link Map map} is provided to allow developer to translate
     * the caption string into map key, e.g. `姓名` to `name`
     *
     * @param sobj          the excel source sobj
     * @param schema        specify the POJO object type
     * @param headerMapping a string-string map that defines caption to map key transform
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(ISObject sobj, Class<T> schema, Map<String, String> headerMapping) {
        return builder().sobject(sobj).headerMapping(headerMapping).build().read(schema);
    }

    public static Builder builder() {
        return new Builder();
    }

    public static Builder builder(Map headerMapping) {
        return new Builder().headerMapping(headerMapping);
    }

    public static Builder builder(TolerantLevel tolerantLevel) {
        return new Builder(tolerantLevel);
    }

    public static Builder builder($.Function<String, String> captionToSchemaTransformer) {
        return new Builder(captionToSchemaTransformer);
    }

    public static Builder builder($.Function<String, String> captionToSchemaTransformer, TolerantLevel tolerantLevel) {
        return new Builder(captionToSchemaTransformer, tolerantLevel);
    }

    private static PushbackInputStream pushbackInputStream(InputStream is) {
        return new PushbackInputStream(is, 100);
    }

    public static class Builder {

        public class HeaderMapper {
            private String caption;

            private HeaderMapper(String caption) {
                E.illegalArgumentIf(S.blank(caption), "caption cannot be null or blank");
                this.caption = caption.trim().toLowerCase();
            }

            public Builder to(String property) {
                E.illegalArgumentIf(S.blank(property), "property cannot be null or blank");
                Builder builder = Builder.this;
                builder.headerMapping.put(caption, property);
                return builder;
            }
        }

        private $.Func0<InputStream> inputStreamProvider;
        private $.Predicate<Sheet> sheetSelector = SheetSelector.ALL;
        // map table header caption to object schema
        private Map<String, String> headerMapping = new HashMap<>();
        private boolean isXlsx = false;
        private int headerRow = 0;
        private boolean ignoreEmptyRows = true;
        private $.Function<String, String> headerTransformer = HeaderTransformStrategy.TO_JAVA_NAME;
        private TolerantLevel tolerantLevel = TolerantLevel.AGGRESSIVE_READ;
        private String terminator;

        public Builder() {
        }

        public Builder(TolerantLevel tolerantLevel) {
            this.tolerantLevel = $.notNull(tolerantLevel);
        }

        public Builder($.Function<String, String> headerTransformer) {
            this.headerTransformer = $.notNull(headerTransformer);
        }

        public Builder($.Function<String, String> headerTransformer, TolerantLevel tolerantLevel) {
            this.headerTransformer = $.notNull(headerTransformer);
            this.tolerantLevel = $.notNull(tolerantLevel);
        }

        public Builder file(final String path) {
            Boolean isXlsx = null;
            if (path.endsWith(".xlsx")) {
                isXlsx = true;
            } else if (path.endsWith(".xls")) {
                isXlsx = false;
            }
            if (null == isXlsx) {
                return inputStream(pushbackInputStream(IO.is(new File(path))));
            }
            this.isXlsx = isXlsx;
            inputStreamProvider = new $.F0<InputStream>() {
                @Override
                public InputStream apply() throws NotAppliedException, Osgl.Break {
                    return new BufferedInputStream(IO.is(new File(path)));
                }
            };
            return this;
        }

        public Builder file(final File file) {
            Boolean isXlsx = null;
            String path = file.getPath();
            if (path.endsWith(".xlsx")) {
                isXlsx = true;
            } else if (path.endsWith(".xls")) {
                isXlsx = false;
            }
            if (null == isXlsx) {
                return inputStream(pushbackInputStream(IO.is(file)));
            }
            this.isXlsx = isXlsx;
            inputStreamProvider = new $.F0<InputStream>() {
                @Override
                public InputStream apply() throws NotAppliedException, Osgl.Break {
                    return new BufferedInputStream(IO.is(file));
                }
            };
            return this;
        }

        public Builder sobject(final ISObject sobj) {
            Boolean isXlsx = null;
            String s = sobj.getAttribute(ISObject.ATTR_FILE_NAME);
            if (S.notBlank(s)) {
                if (s.endsWith(".xlsx")) {
                    isXlsx = true;
                } else if (s.endsWith(".xls")) {
                    isXlsx = false;
                }
            }

            if (null == isXlsx) {
                s = sobj.getAttribute(ISObject.ATTR_CONTENT_TYPE);
                if (S.notBlank(s)) {
                    if (s.startsWith("application/vnd.openxmlformats-officedocument.spreadsheetml")) {
                        isXlsx = true;
                    } else if ("application/vnd.ms-excel".equals(s)) {
                        isXlsx = false;
                    }
                }
                if (null == isXlsx) {
                    return inputStream(pushbackInputStream(sobj.asInputStream()));
                }
            }

            this.isXlsx = isXlsx;
            return sobject(sobj, isXlsx);
        }

        public Builder sobject(final ISObject sobj, boolean isXlsx) {
            this.isXlsx = isXlsx;
            inputStreamProvider = new $.F0<InputStream>() {
                @Override
                public InputStream apply() throws NotAppliedException, Osgl.Break {
                    return sobj.asInputStream();
                }
            };
            return this;
        }

        public Builder classResource(final String url) {
            Boolean xlsx = null;
            if (url.endsWith(".xlsx")) {
                xlsx = true;
            } else if (url.endsWith(".xls")) {
                xlsx = false;
            }
            if (null == xlsx) {
                return inputStream(IO.is(ExcelReader.class.getResource(url)));
            }
            return classResource(url, xlsx);
        }

        public Builder classResource(final String url, boolean isXlsx) {
            this.isXlsx = isXlsx;
            inputStreamProvider = new $.F0<InputStream>() {
                @Override
                public InputStream apply() throws NotAppliedException, Osgl.Break {
                    return IO.is(ExcelReader.class.getResource(url));
                }
            };
            return this;
        }

        public Builder inputStream(final InputStream is) {
            InputStream probeStream = pushbackInputStream(is);
            try {
                return inputStream(probeStream, DocumentFactoryHelper.hasOOXMLHeader(probeStream));
            } catch (IOException e) {
                throw E.ioException(e);
            }
        }

        public Builder inputStream(final InputStream is, boolean isXlsx) {
            this.isXlsx = isXlsx;
            inputStreamProvider = new $.F0<InputStream>() {
                @Override
                public InputStream apply() throws NotAppliedException, Osgl.Break {
                    return is;
                }
            };
            return this;
        }

        public Builder terminator(String terminator) {
            this.terminator = terminator;
            return this;
        }

        public Builder sheetSelector($.Predicate<Sheet> sheetPredicate) {
            sheetSelector = $.notNull(sheetPredicate);
            return this;
        }

        public Builder sheets(String... names) {
            return sheetSelector(SheetSelector.byName(names));
        }

        public Builder excludeSheets(String... names) {
            return sheetSelector(SheetSelector.excludeByName(names));
        }

        public Builder sheets(int... indexes) {
            return sheetSelector(SheetSelector.byPosition(indexes));
        }

        public Builder excludeSheets(int... indexes) {
            return sheetSelector(SheetSelector.excludeByPosition(indexes));
        }

        public Builder headerRow(int rowIndex) {
            E.illegalArgumentIf(rowIndex < 0, "start row must not be negative");
            headerRow = rowIndex;
            return this;
        }

        public Builder ignoreEmptyRows() {
            ignoreEmptyRows = true;
            return this;
        }

        public Builder ignoreEmptyRows(boolean ignore) {
            ignoreEmptyRows = ignore;
            return this;
        }

        public Builder readEmptyRows() {
            ignoreEmptyRows = false;
            return this;
        }

        public HeaderMapper map(String header) {
            return new HeaderMapper(header);
        }

        public Builder headerMapping(Map mapping) {
            E.illegalArgumentIf(mapping.isEmpty(), "empty header mapping found");
            headerMapping = $.cast($.notNull(mapping));
            return this;
        }

        public Builder readColumns(String... headers) {
            return readColumns(C.listOf(headers));
        }


        public Builder readColumns(Collection<String> headers) {
            E.illegalArgumentIf(headers.isEmpty(), "empty read column caption collection found");
            headerMapping = new HashMap<>();
            for (String header : headers) {
                headerMapping.put(header, headerTransformer.apply(header));
            }
            tolerantLevel = TolerantLevel.TOLERANT;
            return this;
        }

        /**
         * Return an {@link ExcelReader} instance from this builder
         *
         * @return the `ExcelReader` built from this builder
         */
        public ExcelReader build() {
            return new ExcelReader(this);
        }
    }
}
