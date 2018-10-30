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

import com.alibaba.fastjson.JSONObject;
import org.junit.Test;
import org.osgl.util.C;
import osgl.ut.TestBase;

import java.util.*;

public class CSVTest extends TestBase {
    @Test
    public void test_no_quote() {
        String line = "10,AU,Australia";
        List<String> result = CsvParser.parseLine(line);

        notNull(result);
        eq(3, result.size());
        eq("10", result.get(0));
        eq("AU", result.get(1));
        eq("Australia", result.get(2));
    }

    @Test
    public void test_no_quote_but_double_quotes_in_column() {
        String line = "10,AU,Aus\"\"tralia";

        List<String> result = CsvParser.parseLine(line);
        notNull(result);
        eq(3, result.size());
        eq("10", result.get(0));
        eq("AU", result.get(1));
        eq("Aus\"tralia", result.get(2));
    }

    @Test
    public void test_double_quotes() {
        String line = "\"10\",\"AU\",\"Australia\"";
        List<String> result = CsvParser.parseLine(line);

        notNull(result);
        eq(3, result.size());
        eq("10", result.get(0));
        eq("AU", result.get(1));
        eq("Australia", result.get(2));
    }

    @Test
    public void test_double_quotes_but_double_quotes_in_column() {
        String line = "\"10\",\"AU\",\"Aus\"\"tralia\"";
        List<String> result = CsvParser.parseLine(line);

        notNull(result);
        eq(3, result.size());
        eq("10", result.get(0));
        eq("AU", result.get(1));
        eq("Aus\"tralia", result.get(2));
    }

    @Test
    public void test_double_quotes_but_comma_in_column() {
        String line = "\"10\",\"AU\",\"Aus,tralia\"";
        List<String> result = CsvParser.parseLine(line);

        notNull(result);
        eq(3, result.size());
        eq("10", result.get(0));
        eq("AU", result.get(1));
        eq("Aus,tralia", result.get(2));
    }

    @Test
    public void testParseWithHeaderRow() {
        List<String> lines = new ArrayList<>();
        lines.add("no,code,country");
        lines.add("10,AU,Australia");
        List<Map<String, String>> data = CsvParser.parse(lines);
        notNull(data);
        eq(1, data.size());
        Map<String, String> row = data.get(0);
        eq("10", row.get("no"));
        eq("AU", row.get("code"));
        eq("Australia", row.get("country"));
    }

    @Test
    public void testParseIntoPojo() {
        List<String> lines = new ArrayList<>();
        lines.add("no,code,country");
        lines.add("10,AU,Australia");
        List<Country> countries = CsvParser.parse(lines, Country.class, C.<String, String>Map("country", "name"));
        eq(1, countries.size());
        Country country = countries.get(0);
        eq(10, country.no);
        eq("AU", country.code);
        eq("Australia", country.name);
    }

    @Test
    public void testParseIntoJSONObject() {
        List<String> lines = new ArrayList<>();
        lines.add("no,code,country");
        lines.add("10,AU,Australia");
        List<JSONObject> list = CsvParser.parse(lines, JSONObject.class, C.<String, String>Map("country", "name"));
        eq(1, list.size());
        JSONObject json = list.get(0);
        eq("10", json.get("no"));
    }
}
