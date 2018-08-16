package org.osgl.xls;

/*-
 * #%L
 * Java Excel Tool
 * %%
 * Copyright (C) 2017 - 2018 OSGL (Open Source General Library)
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
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.osgl.$;
import org.osgl.logging.LogManager;
import org.osgl.logging.Logger;
import org.osgl.util.*;
import osgl.version.Version;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

public class ExcelWriter {

    public static final Version VERSION = ExcelReader.VERSION;
    public static final Logger LOGGER = LogManager.get(ExcelWriter.class);

    private boolean isXlsx;
    private $.Function<String, String> headerTransformer;
    private final Map<String, String> headerMapping;
    private Row headerRow;
    private AtomicInteger maxColId = new AtomicInteger();
    private String dateFormat;
    private String filter;
    private CellStyle dateStyle;
    private CellStyle intStyle;
    private CellStyle doubleStyle;

    public ExcelWriter(boolean isXlsx, String dateFormat, Map<String, String> headerMapping, $.Function<String, String> headerTransformer, String filter) {
        this.isXlsx = isXlsx;
        this.dateFormat = dateFormat;
        this.headerMapping = headerMapping;
        this.filter = filter;
        this.headerTransformer = headerTransformer;
    }

    public void writeSheets(Map<String, Object> data, OutputStream os) {
        Workbook workbook = newWorkbook();
        for (Map.Entry<String, Object> entry : data.entrySet()) {
            Sheet sheet = workbook.createSheet(entry.getKey());
            writeSheet(sheet, entry.getValue());
        }
        commit(workbook, os);
    }

    public void write(Object data, OutputStream os) {
        writeSheet("sheet1", data, os);
    }

    public void write(Object data, File file) {
        isXlsx = file.getName().endsWith(".xlsx");
        write(data, IO.outputStream(file));
    }

    public void writeSheet(String sheetName, Object data, OutputStream os) {
        Workbook workbook = newWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);
        writeSheet(sheet, data);
        commit(workbook, os);
    }

    private void writeSheet(Sheet sheet, Object data) {
        if (null == data) {
            return;
        }
        List list = data instanceof List ? (List) data : C.list(data);
        writeSheet(sheet, list);
    }

    private void writeSheet(Sheet sheet, List<?> list) {
        if (list.isEmpty()) {
            return;
        }
        Object line = firstNonNullValue(list);
        if (null == line) {
            return;
        }
        list = flatMap(list);
        Set<String> headers = extractHeaders((List) list);
        C.Map<String, Integer> colIdx = C.newMap();
        buildColIndex(headers, colIdx);
        createHeaderRow(sheet, colIdx);
        AtomicInteger rowId = new AtomicInteger(1);
        for (Object o: list) {
            if (null == o) {
                continue;
            }
            Map<String, Object> lineData = (Map)o;
            createDataRow(sheet, colIdx, lineData, rowId);
        }
        int max = maxColId.get();
        for (int i = 0; i < max; ++i) {
            sheet.autoSizeColumn(i);
        }
    }

    private Set<String> extractHeaders(List<Map<String, Object>> data) {
        if (S.notBlank(filter)) {
            List<String> list = S.fastSplit(filter, ",");
            if (!list.get(0).startsWith("-")) {
                Set<String> headers = new LinkedHashSet<>();
                headers.addAll(list);
                return headers;
            }
        }
        Set<String> headers = new LinkedHashSet<>();
        for (Map<String, Object> line : data) {
            if (null == line || line.isEmpty()) {
                continue;
            }
            headers.addAll(line.keySet());
        }
        return headers;
    }

    private List flatMap(List<?> list) {
        List retVal = new ArrayList();
        for (Object o : list) {
            if (null == o) {
                continue;
            }
            retVal.add($.flatCopy(o).filter(filter).to(LinkedHashMap.class));
        }
        return retVal;
    }

    private Object firstNonNullValue(List<?> list) {
        for (Object o : list) {
            if (null != o) {
                return o;
            }
        }
        return null;
    }

    private void createDataRow(Sheet sheet, C.Map<String, Integer> colIndex, Map<String, Object> data, AtomicInteger rowId) {
        Row row = sheet.createRow(rowId.getAndIncrement());
        for (Map.Entry<String, Object> cellData : data.entrySet()) {
            int cellId = colIndex.get(cellData.getKey());
            Cell cell = row.createCell(cellId);
            Object val = cellData.getValue();
            if (val instanceof Boolean) {
                cell.setCellValue((Boolean) val);
            } else if (val instanceof Date) {
                cell.setCellValue((Date) val);
                cell.setCellStyle(dateStyle);
            } else if (val instanceof Calendar) {
                cell.setCellValue((Calendar) val);
                cell.setCellStyle(dateStyle);
            } else if (val instanceof Number) {
                cell.setCellValue(((Number) val).doubleValue());
                if (val instanceof Double || val instanceof Float || val instanceof BigDecimal) {
                    cell.setCellStyle(doubleStyle);
                } else {
                    cell.setCellStyle(intStyle);
                }
            } else {
                cell.setCellValue(S.string(val));
            }
        }
    }

    private void createHeaderRow(Sheet sheet, C.Map<String, Integer> colIndex) {
        headerRow = sheet.createRow(0);
        C.Map<Integer, String> flipped = colIndex.flipped();
        int max = maxColId.get();
        for (int i = 0; i < max; ++i) {
            Cell cell = headerRow.createCell(i);
            String label = flipped.get(i);
            boolean mapped = false;
            if (null != headerMapping) {
                String newLabel = headerMapping.get(label);
                if (null != newLabel) {
                    label = newLabel;
                    mapped = true;
                }
            }
            if (!mapped && null != headerTransformer) {
                label = headerTransformer.apply(label);
            }
            cell.setCellValue(label);
        }
    }

    private void buildColIndex(Collection<String> keys, C.Map<String, Integer> colIndex) {
        for (String key : keys) {
            if (colIndex.containsKey(key)) {
                continue;
            }
            colIndex.put(key, maxColId.getAndIncrement());
        }
    }

    private void commit(Workbook workbook, OutputStream os) {
        try {
            workbook.write(os);
        } catch (IOException e) {
            throw E.ioException(e);
        } finally {
            IO.close(os);
        }
    }

    private Workbook newWorkbook() {
        Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat(dateFormat));
        intStyle = workbook.createCellStyle();
        intStyle.setDataFormat(createHelper.createDataFormat().getFormat("0"));
        doubleStyle = workbook.createCellStyle();
        doubleStyle.setDataFormat(createHelper.createDataFormat().getFormat("0.00"));
        return workbook;
    }

    public static class Builder {
        private boolean isXlsx;
        private String dateFormat;
        private Map<String, String> headerMapping;
        private String filter;
        private $.Function<String, String> headerTransformer;

        public Builder asXlsx() {
            this.isXlsx = true;
            return this;
        }

        public Builder filter(String filter) {
            this.filter = filter;
            return this;
        }

        public Builder dateFormat(String dateFormat) {
            this.dateFormat = dateFormat;
            return this;
        }

        public Builder headerTransformer($.Function<String, String> headerTransformer) {
            this.headerTransformer = headerTransformer;
            return this;
        }

        public Builder headerMap(Map<String, String> headerMap) {
            this.headerMapping = headerMap;
            return this;
        }

        public ExcelWriter build() {
            return new ExcelWriter(isXlsx, dateFormat, headerMapping, headerTransformer, filter);
        }
    }

    public static Builder builder() {
        return new Builder();
    }
}
