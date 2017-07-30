package org.osgl.xls;

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

import java.io.*;
import java.util.*;

public class ExcelReader {

    public static final Logger LOGGER = LogManager.get(ExcelReader.class);

    private final boolean isXlsx;
    private final $.Func0<InputStream> inputStreamProvider;
    private final $.Predicate<Sheet> sheetSelector;
    private final int captionRow;
    private final boolean ignoreEmptyRows;
    private final TolerantLevel tolerantLevel;
    private final $.Function<String, String> captionToSchemaTransformer;
    private final Map<String, String> captionSchemaMapping;

    private ExcelReader(Builder builder) {
        inputStreamProvider = $.notNull(builder.inputStreamProvider);
        sheetSelector = builder.sheetSelector;
        captionRow = builder.captionRow;
        ignoreEmptyRows = builder.ignoreEmptyRows;
        isXlsx = builder.isXlsx;
        tolerantLevel = builder.tolerantLevel;
        captionSchemaMapping = builder.captionSchemaMapping;
        captionToSchemaTransformer = builder.captionToSchemaTransformer;
    }

    public List<Map<String, Object>> read() {
        return (List)read(Map.class);
    }

    public <TYPE> List<TYPE> read(Class<? extends TYPE> schema) {
        final List<TYPE> dataList = new ArrayList<>();
        final Workbook wb = loadWorkbook();
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
            Map<String, PropertySetter> pojoSetters = JavaBeanUtil.setters(schema);
            schemaMapping.putAll(pojoSetters);
            for (Map.Entry<String, String> entry : captionSchemaMapping.entrySet()) {
                String caption = entry.getKey().trim().toLowerCase();
                String field = entry.getValue();
                PropertySetter setter = pojoSetters.get(field);
                if (null != setter) {
                    schemaMapping.put(caption, setter);
                } else if (tolerantLevel.isTolerant()) {
                    field = captionToSchemaTransformer.apply(caption);
                    setter = pojoSetters.get(field);
                    if (null != setter) {
                        schemaMapping.put(caption, setter);
                    }
                }
            }
        } else {
            for (Map.Entry<String, String> entry : captionSchemaMapping.entrySet()) {
                schemaMapping.put(entry.getKey(), new MapSetter(entry.getValue()));
            }
        }
        return schemaMapping;
    }

    private <TYPE> void read(Workbook workbook, final List<TYPE> dataList, Map<String, PropertySetter> setterMap, Class<? extends TYPE>  schema) {
        for (Sheet sheet : workbook) {
            if (!sheetSelector.test(sheet)) {
                continue;
            }
            read(sheet, dataList, setterMap, schema);
        }
    }

    private <TYPE> void read(Sheet sheet, final List<TYPE> dataList, Map<String, PropertySetter> setterMap, Class<? extends TYPE>  schema) {
        $.Var<Integer> captionRowHolder = $.var(this.captionRow);
        boolean schemaIsMap = Map.class.isAssignableFrom(schema);
        Map<Integer, PropertySetter> columnIndex = buildColumnIndex(sheet, setterMap, schemaIsMap, captionRowHolder);
        if (columnIndex.size() < setterMap.size()) {
            tolerantLevel.columnIndexMapNotFullyBuilt(sheet);
        }
        if (columnIndex.isEmpty()) {
            return;
        }
        int startRow = captionRowHolder.get() + 1;
        int maxRow = sheet.getLastRowNum() + 1;
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
            if (isEmptyRow && !ignoreEmptyRows) {
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

    private Map<Integer, PropertySetter> buildColumnIndex(Sheet sheet, Map<String, PropertySetter> setterMap, boolean schemaIsMap, $.Var<Integer> captionRowHolder) {
        int startRow = sheet.getFirstRowNum();
        int maxRow = sheet.getLastRowNum();
        int captionRow = this.captionRow;
        if (captionRow < startRow || captionRow >= maxRow) {
            tolerantLevel.captionRowOutOfScope(sheet);
            captionRow = startRow;
        }
        if (tolerantLevel.isStrict()) {
            return buildColumnIndex(sheet.getRow(captionRow), setterMap, schemaIsMap);
        }
        Map<Integer, PropertySetter> index = C.map();
        for (int rowId = captionRow; rowId < maxRow; ++rowId) {
            index = buildColumnIndex(sheet.getRow(rowId), setterMap, schemaIsMap);
            if (!index.isEmpty()) {
                captionRowHolder.set(rowId);
                return index;
            }
        }
        return index;
    }

    private Map<Integer, PropertySetter> buildColumnIndex(Row row, Map<String, PropertySetter> setterMap, boolean schemaIsMap) {
        Map<Integer, PropertySetter> retVal = new HashMap<>();
        for (Cell cell : row) {
            try {
                String caption = cell.getStringCellValue();
                if (S.blank(caption)) {
                    continue;
                }
                caption = caption.trim();
                PropertySetter setter = null;
                String translated = captionSchemaMapping.get(caption.toLowerCase());
                if (null != translated) {
                    setter = setterMap.get(translated);
                }
                if (null == setter) {
                    String key = captionToSchemaTransformer.apply(caption);
                    setter = setterMap.get(key);
                }
                if (null != setter) {
                    retVal.put(cell.getColumnIndex(), setter);
                } else if (tolerantLevel.isAggressiveReading() && schemaIsMap) {
                    retVal.put(cell.getColumnIndex(), new MapSetter(captionToSchemaTransformer.apply(caption)));
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

        public void captionRowOutOfScope(Sheet sheet) {
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
     * A `captionSchemaMapping` {@link Map map} is provided to allow developer to translate
     * the caption string into map key, e.g. `姓名` to `name`
     *
     * @param file the excel source file
     * @param captionSchemaMapping a string-string map that defines caption to map key transform
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(File file, Map<String, String> captionSchemaMapping) {
        return builder().file(file).captionSchemaMapping(captionSchemaMapping).build().read();
    }

    /**
     * A quick API for reading an excel file and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * @param file the excel source file
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
     * A `captionSchemaMapping` {@link Map map} is provided to allow developer to translate
     * the caption string into map key, e.g. `姓名` to `name`
     *
     * @param file the excel source file
     * @param schema specify the POJO object type
     * @param captionSchemaMapping a string-string map that defines caption to map key transform
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(File file, Class<T> schema, Map<String, String> captionSchemaMapping) {
        return builder().file(file).captionSchemaMapping(captionSchemaMapping).build().read(schema);
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
     * A `captionSchemaMapping` {@link Map map} is provided to allow developer to translate
     * the caption string into map key, e.g. `姓名` to `name`
     *
     * @param inputStream the excel source inputStream
     * @param captionSchemaMapping a string-string map that defines caption to map key transform
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(InputStream inputStream, Map<String, String> captionSchemaMapping) {
        return builder().inputStream(inputStream).captionSchemaMapping(captionSchemaMapping).build().read();
    }

    /**
     * A quick API for reading an excel inputStream and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * @param inputStream the excel source inputStream
     * @param schema specify the POJO object type
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
     * A `captionSchemaMapping` {@link Map map} is provided to allow developer to translate
     * the caption string into map key, e.g. `姓名` to `name`
     *
     * @param inputStream the excel source inputStream
     * @param schema specify the POJO object type
     * @param captionSchemaMapping a string-string map that defines caption to map key transform
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(InputStream inputStream, Class<T> schema, Map<String, String> captionSchemaMapping) {
        return builder().inputStream(inputStream).captionSchemaMapping(captionSchemaMapping).build().read(schema);
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
     * A `captionSchemaMapping` {@link Map map} is provided to allow developer to translate
     * the caption string into map key, e.g. `姓名` to `name`
     *
     * @param sobj the excel source sobj
     * @param captionSchemaMapping a string-string map that defines caption to map key transform
     * @return the map list as described above
     */
    public static List<Map<String, Object>> read(ISObject sobj, Map<String, String> captionSchemaMapping) {
        return builder().sobject(sobj).captionSchemaMapping(captionSchemaMapping).build().read();
    }

    /**
     * A quick API for reading an excel sobj and return a {@link List list} of POJO object
     * instances where the type is specified by `schema` parameter, where each pojo
     * instance is corresponding to an excel sheet row
     *
     * @param sobj the excel source sobj
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
     * A `captionSchemaMapping` {@link Map map} is provided to allow developer to translate
     * the caption string into map key, e.g. `姓名` to `name`
     *
     * @param sobj the excel source sobj
     * @param schema specify the POJO object type
     * @param captionSchemaMapping a string-string map that defines caption to map key transform
     * @return the pojo object list as described above
     */
    public static <T> List<T> read(ISObject sobj, Class<T> schema, Map<String, String> captionSchemaMapping) {
        return builder().sobject(sobj).captionSchemaMapping(captionSchemaMapping).build().read(schema);
    }

    public static Builder builder() {
        return new Builder();
    }

    public static Builder builder(Map captionSchemaMapping) {
        return new Builder().captionSchemaMapping(captionSchemaMapping);
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

        public class ColumMapper {
            private String caption;
            private ColumMapper(String caption) {
                E.illegalArgumentIf(S.blank(caption), "caption cannot be null or blank");
                this.caption = caption.trim().toLowerCase();
            }
            public Builder to(String property) {
                E.illegalArgumentIf(S.blank(property), "property cannot be null or blank");
                Builder builder = Builder.this;
                builder.captionSchemaMapping.put(caption, property);
                return builder;
            }
        }

        private $.Func0<InputStream> inputStreamProvider;
        private $.Predicate<Sheet> sheetSelector = SheetSelector.ALL;
        // map table header caption to object schema
        private Map<String, String> captionSchemaMapping = new HashMap<>();
        private boolean isXlsx = false;
        private int captionRow = 0;
        private boolean ignoreEmptyRows = true;
        private $.Function<String, String> captionToSchemaTransformer = CaptionSchemaTransformStrategy.TO_JAVA_NAME;
        private TolerantLevel tolerantLevel = TolerantLevel.AGGRESSIVE_READ;

        public Builder() {
        }

        public Builder(TolerantLevel tolerantLevel) {
            this.tolerantLevel = $.notNull(tolerantLevel);
        }

        public Builder($.Function<String, String> captionToSchemaTransformer) {
            this.captionToSchemaTransformer = $.notNull(captionToSchemaTransformer);
        }

        public Builder($.Function<String, String> captionToSchemaTransformer, TolerantLevel tolerantLevel) {
            this.captionToSchemaTransformer = $.notNull(captionToSchemaTransformer);
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

        public Builder sheetSelector($.Predicate<Sheet> sheetPredicate) {
            sheetSelector = $.notNull(sheetPredicate);
            return this;
        }

        public Builder sheets(String ... names) {
            return sheetSelector(SheetSelector.byName(names));
        }

        public Builder excludeSheets(String... names) {
            return sheetSelector(SheetSelector.excludeByName(names));
        }

        public Builder sheets(int ... indexes) {
            return sheetSelector(SheetSelector.byPosition(indexes));
        }

        public Builder excludeSheets(int... indexes) {
            return sheetSelector(SheetSelector.excludeByPosition(indexes));
        }

        public Builder captionRow(int rowIndex) {
            E.illegalArgumentIf(rowIndex < 0, "start row must not be negative");
            captionRow = rowIndex;
            return this;
        }

        public Builder ignoreEmptyRows() {
            ignoreEmptyRows = true;
            return this;
        }

        public Builder readEmptyRows() {
            ignoreEmptyRows = false;
            return this;
        }

        public ColumMapper map(String caption) {
            return new ColumMapper(caption);
        }

        public Builder captionSchemaMapping(Map mapping) {
            E.illegalArgumentIf(mapping.isEmpty(), "empty caption to schema mapping found");
            captionSchemaMapping = $.cast($.notNull(mapping));
            return this;
        }

        public Builder readColumns(String... captions) {
            return readColumns(C.listOf(captions));
        }

        public Builder readColumns(Collection<String> captions) {
            E.illegalArgumentIf(captions.isEmpty(), "empty read column caption collection found");
            captionSchemaMapping = new HashMap<>();
            for (String caption : captions) {
                captionSchemaMapping.put(caption, captionToSchemaTransformer.apply(caption));
            }
            return this;
        }

        public ExcelReader build() {
            return new ExcelReader(this);
        }
    }
}
