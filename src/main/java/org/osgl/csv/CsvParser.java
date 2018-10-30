package org.osgl.csv;

/*-
 * #%L
 * Java CSV Tool
 * %%
 * Copyright (C) 2018 OSGL (Open Source General Library)
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

import org.osgl.$;
import org.osgl.util.*;

import java.util.*;

/**
 * CSV parser
 *
 * The code come from https://www.mkyong.com/java/how-to-read-and-parse-csv-file-in-java/
 */
public class CsvParser {

    public static final char DEFAULT_QUOTE = '"';

    public static final char DEFAULT_SEPARATOR = ',';

    static {
        register();
    }


    public static void register() {
        IO.registerInputStreamHandler(MimeType.Trait.csv, new CsvInputStreamHandler());
    }

    /**
     * Parse a list of lines into list of LinkedHashMap.
     *
     * **Note** the first line must be the header line
     *
     * @param lines
     *      the csv lines
     * @return a list of LinkedHashMap
     */
    public static List<Map<String, String>> parse(List<String> lines) {
        if (lines.isEmpty()) {
            return new ArrayList<>();
        }
        String headerLine = lines.remove(0);
        return parse(headerLine, lines);
    }

    /**
     * Parse a list of lines into list of LinkedHashMap.
     *
     * **Note** the line list must not contains header line
     *
     * @param headerLine
     *      the header line
     * @param lines
     *      the csv lines
     * @return a list of LinkedHashMap
     */
    public static List<Map<String, String>> parse(String headerLine, List<String> lines) {
        int totalLines = lines.size();
        if (0 >= totalLines) {
            return new ArrayList<>();
        }
        List<String> headers = parseLine(headerLine);
        int columnCnt = headers.size();
        String[] headerLookup = new String[columnCnt];
        headerLookup = headers.toArray(headerLookup);
        List<Map<String, String>> retVal = new ArrayList<>();
        for (int i = 0; i < totalLines; ++i) {
            retVal.add(buildRecord(columnCnt, headerLookup, parseLine(lines.get(i))));
        }
        return retVal;
    }

    public static Map<String, String> parseDataLine(String headerLine, String dataLine) {
        List<Map<String, String>> list = parse(headerLine, C.list(dataLine));
        return list.get(0);
    }

    /**
     * Parse a list of lines into list of recordType specified.
     *
     * **Note** the first line must be the header line
     *
     * @param lines
     *      the csv lines
     * @param recordType
     *      the record type
     * @return a list of record type
     */
    public static <T> List<T> parse(List<String> lines, Class<T> recordType) {
        return parse(lines, recordType, null);
    }

    /**
     * Parse a list of lines into list of recordType specified.
     *
     * **Note** the line list must not contains header line
     *
     * @param header
     *      the header line
     * @param lines
     *      the csv lines
     * @param recordType
     *      the record type
     * @return a list of record type
     */
    public static <T> List<T> parse(String header, List<String> lines, Class<T> recordType) {
        return parse(header, lines, recordType, null);
    }

    public static <T> T parseDataLine(String header, String dataLine, Class<T> recordType) {
        return parseDataLine(header, dataLine, recordType, null);
    }

    /**
     * Parse a list of lines into list of recordType specified.
     *
     * **Note** the first line must be the header line
     *
     * @param lines
     *      the csv lines
     * @param recordType
     *      the record type
     * @param headerMapping
     *      the header mapping
     * @return a list of record type
     */
    public static <T> List<T> parse(List<String> lines, Class<T> recordType, Map<String, String> headerMapping) {
        if (lines.isEmpty()) {
            return new ArrayList<>();
        }
        String header = lines.remove(0);
        return parse(header, lines, recordType, headerMapping);
    }

    /**
     * Parse a list of lines into list of recordType specified.
     *
     * **Note** the line list must not contains header line
     *
     * @param header
     *      the header line
     * @param lines
     *      the csv lines
     * @param recordType
     *      the record type
     * @param headerMapping
     *      the header mapping
     * @return a list of record type
     */
    public static <T> List<T> parse(String header, List<String> lines, Class<T> recordType, Map<String, String> headerMapping) {
        int totalLines = lines.size();
        if (0 >= totalLines) {
            return new ArrayList<>();
        }
        List<String> headerLine = parseLine(header);
        int columnCnt = headerLine.size();
        String[] headerLookup = new String[columnCnt];
        headerLookup = headerLine.toArray(headerLookup);
        List<T> retVal = new ArrayList<>();
        for (int i = 0; i < totalLines; ++i) {
            Map<String, String> row = buildRecord(columnCnt, headerLookup, parseLine(lines.get(i)));
            if (null != headerMapping) {
                retVal.add($.map(row).withHeadMapping(headerMapping).to(recordType));
            } else {
                retVal.add($.map(row).to(recordType));
            }
        }
        return retVal;
    }

    public static <T> T parseDataLine(String header, String line, Class<T> recordType, Map<String, String> headerMapping) {
        List<String> headerLine = parseLine(header);
        int columnCnt = headerLine.size();
        String[] headerLookup = new String[columnCnt];
        headerLookup = headerLine.toArray(headerLookup);
        Map<String, String> row = buildRecord(columnCnt, headerLookup, parseLine(line));
        return null != headerMapping ? $.map(row).withHeadMapping(headerMapping).to(recordType) : $.map(row).to(recordType);
    }

    private static LinkedHashMap<String, String> buildRecord(int columnCnt, String[] headerLookup, List<String> row) {
        LinkedHashMap<String, String> record = new LinkedHashMap<>();
        if (row.isEmpty()) {
            return record;
        }
        for (int i = 0; i < columnCnt; i++) {
            record.put(headerLookup[i], row.get(i));
        }
        return record;
    }

    public static List<String> parseLine(String line) {
        return parseLine(line, ' ', ' ');
    }

    public static List<String> parseLine(String line, char separator, char quote) {

        List<String> result = new ArrayList<>();

        //if empty, return!
        if (line == null || line.isEmpty()) {
            return result;
        }

        if (quote == ' ') {
            quote = DEFAULT_QUOTE;
        }

        if (separator == ' ') {
            separator = DEFAULT_SEPARATOR;
        }

        StringBuffer curVal = new StringBuffer();
        boolean inQuotes = false;
        boolean startCollectChar = false;
        boolean doubleQuotesInColumn = false;

        char[] chars = line.toCharArray();

        for (char ch : chars) {

            if (inQuotes) {
                startCollectChar = true;
                if (ch == quote) {
                    inQuotes = false;
                    doubleQuotesInColumn = false;
                } else {

                    //Fixed : allow "" in custom quote enclosed
                    if (ch == '\"') {
                        if (!doubleQuotesInColumn) {
                            curVal.append(ch);
                            doubleQuotesInColumn = true;
                        }
                    } else {
                        curVal.append(ch);
                    }

                }
            } else {
                if (ch == quote) {

                    inQuotes = true;

                    //Fixed : allow "" in empty quote enclosed
                    if (chars[0] != '"' && quote == '\"') {
                        curVal.append('"');
                    }

                    //double quotes in column will hit this!
                    if (startCollectChar) {
                        curVal.append('"');
                    }

                } else if (ch == separator) {

                    result.add(curVal.toString());

                    curVal = new StringBuffer();
                    startCollectChar = false;

                } else if (ch == '\r') {
                    //ignore LF characters
                    continue;
                } else if (ch == '\n') {
                    //the end, break!
                    break;
                } else {
                    curVal.append(ch);
                }
            }

        }

        result.add(curVal.toString().trim());

        return result;
    }

}
