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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.osgl.$;
import org.osgl.logging.LogManager;
import org.osgl.logging.Logger;
import org.osgl.util.C;
import org.osgl.util.E;
import org.osgl.util.IO;
import org.osgl.util.S;
import osgl.version.Version;
import osgl.version.Versioned;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

@Versioned
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
    private Map<String, String> fieldStylePatterns;
    private boolean bigData;
    private boolean freezeTopRow = true;
    private boolean autoFilter = true;
    private SheetStyle sheetStyle;

    private CellStyle topRowStyle;
    private CellStyle dataRowStyle;
    private CellStyle alternateDataRowStyle;

    private CellStyle dataRowIntStyle;
    private CellStyle alternateDataRowIntStyle;

    private CellStyle dataRowDateStyle;
    private CellStyle alternateDataRowDateStyle;

    private CellStyle dataRowDoubleStyle;
    private CellStyle alternateDataRowDoubleStyle;
    private Map<$.T2<CellStyle, String>, CellStyle> specialStyleMap = new HashMap<>();

    public ExcelWriter(boolean isXlsx, String dateFormat, Map<String, String> headerMapping,
                       Map<String, String> fieldStylePatterns, $.Function<String, String> headerTransformer,
                       String filter, boolean bigData, SheetStyle sheetStyle) {
        this.isXlsx = isXlsx;
        this.dateFormat = dateFormat;
        this.headerMapping = headerMapping;
        this.fieldStylePatterns = null != fieldStylePatterns ? fieldStylePatterns : C.<String, String>Map();
        this.filter = filter;
        this.headerTransformer = headerTransformer;
        this.bigData = bigData;
        this.sheetStyle = sheetStyle;
        if (null == this.sheetStyle) {
            this.sheetStyle = SheetStyleManager.getDefault();
        }
    }

    public ExcelWriter(boolean isXlsx, String dateFormat, Map<String, String> headerMapping,
                       Map<String, String> fieldStylePatterns, $.Function<String, String> headerTransformer,
                       String filter, boolean bigData, String sheetStyleId) {
        this(isXlsx, dateFormat, headerMapping, fieldStylePatterns, headerTransformer, filter, bigData, SheetStyleManager.SINGLETON.getSheetStyle(sheetStyleId));
    }

    public ExcelWriter noFreeTopRow() {
        freezeTopRow = false;
        return this;
    }

    public ExcelWriter noAutoFilter() {
        autoFilter = false;
        return this;
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
        if (data instanceof Map) {
            writeSheets((Map<String, Object>) data, os);
        } else {
            writeSheet("sheet1", data, os);
        }
    }

    public void writeSheets(Map<String, Object> data, File file) {
        write(data, IO.outputStream(file));
    }

    public void write(Object data, File file) {
        if (data instanceof Map) {
            writeSheets((Map<String, Object>) data, file);
        } else {
            isXlsx = file.getName().endsWith(".xlsx");
            write(data, IO.outputStream(file));
        }
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
        if (null != sheetStyle) {
            sheet.setDisplayGridlines(sheetStyle.displayGridLine);
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
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        }
        buildColIndex(headers, colIdx);
        createHeaderRow(sheet, colIdx);
        AtomicInteger rowId = new AtomicInteger(1);
        int maxCol = maxColId.get();
        for (Object o : list) {
            if (null == o) {
                continue;
            }
            Map<String, Object> lineData = (Map) o;
            createDataRow(sheet, colIdx, lineData, rowId, maxCol);
        }
        int max = maxColId.get();
        for (int i = 0; i < max; ++i) {
            sheet.autoSizeColumn(i);
            int w = sheet.getColumnWidth(i);
            if (w / 256 > 50) {
                sheet.setColumnWidth(i, 50 * 256);
            }
        }
        if (freezeTopRow) {
            sheet.createFreezePane(0, 1);
        }
        if (autoFilter) {
            sheet.setAutoFilter(new CellRangeAddress(headerRow.getRowNum(), headerRow.getRowNum(), headerRow.getFirstCellNum(), headerRow.getLastCellNum() - 1));
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

    private void createDataRow(Sheet sheet, C.Map<String, Integer> colIndex, Map<String, Object> data, AtomicInteger rowId, int maxCol) {
        int rowNumber = rowId.getAndIncrement();
        Row row = sheet.createRow(rowNumber);
        boolean isAlternate = rowNumber % 2 == 0;
        CellStyle cellStyle = isAlternate ? alternateDataRowStyle : dataRowStyle;
        //row.setRowStyle(cellStyle);
        Set<Integer> cols = new HashSet<>(maxCol);
        for (int i = 0; i < maxCol; ++i) {
            cols.add(i);
        }
        for (Map.Entry<String, Object> cellData : data.entrySet()) {
            String key = cellData.getKey();
            Object val = cellData.getValue();
            String stylePattern = fieldStylePatterns.get(key);
            int cellId = colIndex.get(key);
            cols.remove(cellId);
            writeCell(sheet, row, cellId, stylePattern, cellStyle, val, isAlternate);
        }
        for (int cellId : cols) {
            writeCell(sheet, row, cellId, null, cellStyle, "", isAlternate);
        }
    }

    private void writeCell(Sheet sheet, Row row, int cellId, String stylePattern, CellStyle cellStyle, Object val, boolean isAlternate) {
        Cell cell = row.createCell(cellId);
        CellStyle specialStyle = null;
        if (null != stylePattern) {
            specialStyle = getSpecialFormatStyle(sheet.getWorkbook(), cellStyle, stylePattern);
        }
        if (val instanceof Boolean) {
            cell.setCellValue((Boolean) val);
            cell.setCellStyle(null == specialStyle ? cellStyle : specialStyle);
        } else if (val instanceof Date) {
            cell.setCellValue((Date) val);
            cell.setCellStyle(null == specialStyle ? (isAlternate ? alternateDataRowDateStyle : dataRowDateStyle) : specialStyle);
        } else if (val instanceof Calendar) {
            cell.setCellValue((Calendar) val);
            cell.setCellStyle(null == specialStyle ? (isAlternate ? alternateDataRowDateStyle : dataRowDateStyle) : specialStyle);
        } else if (val instanceof Number) {
            cell.setCellValue(((Number) val).doubleValue());
            if (val instanceof Double || val instanceof Float || val instanceof BigDecimal) {
                cell.setCellStyle(null == specialStyle ? (isAlternate ? alternateDataRowDoubleStyle : dataRowDoubleStyle) : specialStyle);
            } else {
                cell.setCellStyle(null == specialStyle ? (isAlternate ? alternateDataRowIntStyle : dataRowIntStyle) : specialStyle);
            }
        } else {
            String s = S.string(val);
            if (S.isEmpty(s)) {
                s = " "; // force it fill background color
            }
            cell.setCellValue(s);
            cell.setCellStyle(null == specialStyle ? cellStyle : specialStyle);
        }

    }

    private void createHeaderRow(Sheet sheet, C.Map<String, Integer> colIndex) {
        headerRow = sheet.createRow(0);
        C.Map<Integer, String> flipped = colIndex.flipped();
        int max = maxColId.get();
        CellStyle cellStyle = topRowStyle;
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
            cell.setCellStyle(cellStyle);
        }
    }

    private static CellStyle createStyle(Workbook workbook, RowStyle style, RowStyle parentStyle, boolean boldFont) {
        if (null == style) {
            return null;
        }
        CellStyle cellStyle = workbook.createCellStyle();
        BorderStyle borderStyle = style.topBorderStyle;
        if (null == borderStyle) {
            borderStyle = parentStyle.topBorderStyle;
        }
        if (null != borderStyle) {
            cellStyle.setBorderTop(borderStyle);
        }
        borderStyle = style.rightBorderStyle;
        if (null == borderStyle) {
            borderStyle = parentStyle.rightBorderStyle;
        }
        if (null != borderStyle) {
            cellStyle.setBorderRight(borderStyle);
        }
        borderStyle = style.bottomBorderStyle;
        if (null == borderStyle) {
            borderStyle = parentStyle.bottomBorderStyle;
        }
        if (null != borderStyle) {
            cellStyle.setBorderBottom(borderStyle);
        }
        borderStyle = style.leftBorderStyle;
        if (null == borderStyle) {
            borderStyle = parentStyle.leftBorderStyle;
        }
        if (null != borderStyle) {
            cellStyle.setBorderLeft(borderStyle);
        }
        IndexedColors borderColor = style.topBorderColor;
        if (null == borderColor) {
            borderColor = parentStyle.topBorderColor;
        }
        if (null != borderColor) {
            cellStyle.setTopBorderColor(borderColor.getIndex());
        }
        borderColor = style.rightBorderColor;
        if (null == borderColor) {
            borderColor = parentStyle.rightBorderColor;
        }
        if (null != borderColor) {
            cellStyle.setRightBorderColor(borderColor.getIndex());
        }
        borderColor = style.bottomBorderColor;
        if (null == borderColor) {
            borderColor = parentStyle.bottomBorderColor;
        }
        if (null != borderColor) {
            cellStyle.setBottomBorderColor(borderColor.getIndex());
        }
        borderColor = style.leftBorderColor;
        if (null == borderColor) {
            borderColor = parentStyle.leftBorderColor;
        }
        if (null != borderColor) {
            cellStyle.setLeftBorderColor(borderColor.getIndex());
        }
        IndexedColors bgColor = style.bgColor;
        if (null == bgColor) {
            bgColor = parentStyle.bgColor;
        }
        if (null != bgColor) {
            cellStyle.setFillForegroundColor(bgColor.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        IndexedColors fgColor = style.fgColor;
        if (null == fgColor) {
            fgColor = parentStyle.fgColor;
        }
        Font font = workbook.createFont();
        if (boldFont) {
            font.setBold(true);
        }
        if (null != fgColor) {
            font.setColor(fgColor.getIndex());
        }
        cellStyle.setFont(font);
        return cellStyle;
    }

    private void buildColIndex(Collection<String> keys, C.Map<String, Integer> colIndex) {
        maxColId.set(0);
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
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }
            IO.close(os);
        }
    }

    private CellStyle createIntStyle(Workbook workbook, CellStyle cloneFrom) {
        return createSpecialFormatStyle(workbook, cloneFrom, "0");
    }

    private CellStyle createDoubleStyle(Workbook workbook, CellStyle cloneFrom) {
        return createSpecialFormatStyle(workbook, cloneFrom, "0.00");
    }

    private CellStyle createDateStyle(Workbook workbook, CellStyle cloneFrom) {
        return createSpecialFormatStyle(workbook, cloneFrom, dateFormat);
    }

    private CellStyle getSpecialFormatStyle(Workbook workbook, CellStyle cloneFrom, String format) {
        $.T2<CellStyle, String> key = $.T2(cloneFrom, format);
        CellStyle style = specialStyleMap.get(key);
        if (null == style) {
            style = createSpecialFormatStyle(workbook, cloneFrom, format);
            specialStyleMap.put(key, style);
        }
        return style;
    }

    private CellStyle createSpecialFormatStyle(Workbook workbook, CellStyle cloneFrom, String format) {
        CreationHelper createHelper = workbook.getCreationHelper();
        CellStyle style = workbook.createCellStyle();
        if (null != cloneFrom) {
            style.cloneStyleFrom(cloneFrom);
        }
        style.setDataFormat(createHelper.createDataFormat().getFormat(format));
        return style;
    }

    private Workbook newWorkbook() {
        Workbook workbook = isXlsx ? (bigData ? new SXSSFWorkbook() : new XSSFWorkbook()) : new HSSFWorkbook();
        if (null != sheetStyle) {
            topRowStyle = createStyle(workbook, sheetStyle.topRowStyle, sheetStyle.topRowStyle, true);
            dataRowStyle = createStyle(workbook, sheetStyle.dataRowStyle, sheetStyle.dataRowStyle, false);
            alternateDataRowStyle = createStyle(workbook, sheetStyle.alternateDataRowStyle, sheetStyle.dataRowStyle, false);
            if (null == alternateDataRowStyle) {
                alternateDataRowStyle = dataRowStyle;
            }
        }

        dataRowDateStyle = createDateStyle(workbook, dataRowStyle);
        dataRowIntStyle = createIntStyle(workbook, dataRowStyle);
        dataRowDoubleStyle = createDoubleStyle(workbook, dataRowStyle);

        alternateDataRowDateStyle = createDateStyle(workbook, alternateDataRowStyle);
        alternateDataRowIntStyle = createIntStyle(workbook, alternateDataRowStyle);
        alternateDataRowDoubleStyle = createDoubleStyle(workbook, alternateDataRowStyle);

        return workbook;
    }

    public static class Builder {
        private boolean isXlsx;
        private String dateFormat;
        private Map<String, String> headerMapping = new HashMap<>();
        private Map<String, String> fieldStylePatterns = new HashMap<>();
        private String filter;
        private boolean bigData;
        private SheetStyle sheetStyle;
        private $.Function<String, String> headerTransformer;

        public Builder asXlsx() {
            this.isXlsx = true;
            return this;
        }

        public Builder filter(String filter) {
            this.filter = filter;
            return this;
        }

        public Builder sheetStyle(SheetStyle sheetStyle) {
            this.sheetStyle = sheetStyle;
            return this;
        }

        public Builder sheetStyle(String sheetStyleId) {
            this.sheetStyle = SheetStyleManager.SINGLETON.getSheetStyle(sheetStyleId);
            return this;
        }

        public Builder dateFormat(String dateFormat) {
            this.dateFormat = dateFormat;
            return this;
        }

        public Builder bigData() {
            this.bigData = true;
            return this;
        }

        public Builder fieldStylePatterns(Map<String, String> patterns) {
            this.fieldStylePatterns.putAll(patterns);
            return this;
        }

        public Builder fieldStylePattern(String fieldName, String style) {
            this.fieldStylePatterns.put(fieldName, style);
            return this;
        }

        public Builder headerTransformer($.Function<String, String> headerTransformer) {
            this.headerTransformer = headerTransformer;
            return this;
        }

        public Builder headerMap(Map<String, String> headerMap) {
            this.headerMapping.putAll(headerMap);
            return this;
        }

        public Builder mapHeader(String fieldName, String header) {
            this.headerMapping.put(fieldName, header);
            return this;
        }

        public ExcelWriter build() {
            return new ExcelWriter(isXlsx, dateFormat, headerMapping, fieldStylePatterns, headerTransformer, filter, bigData, sheetStyle);
        }
    }

    public static Builder builder() {
        return new Builder();
    }
}
