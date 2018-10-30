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

import org.osgl.util.C;
import org.osgl.util.E;

import java.io.Writer;
import java.util.List;
import java.util.Map;

public class CsvWriter {

    public static <T> void write(Iterable<T> data, Class<T> componentType, Writer writer) {
        if (Map.class.isAssignableFrom(componentType)) {
            write((Iterable<Map<String, Object>>) data, writer);
        } else {

        }
    }

    public static void write(Iterable<Map<String, Object>> data, Writer writer) {
        boolean first = true;
        List<String> keys = null;
        for (Map<String, Object> row : data) {
            if (first) {
                keys = C.list(row.keySet());
            }
        }
    }

    private static void writeHead(List<String> keys, Writer writer) {
        if (keys.isEmpty()) {
            return;
        }
        E.tbd();
    }


}
