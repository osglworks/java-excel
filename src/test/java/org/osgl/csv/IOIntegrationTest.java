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

import org.junit.Test;
import org.osgl.util.*;
import osgl.ut.TestBase;

import java.net.URL;
import java.util.List;
import java.util.Map;

public class IOIntegrationTest extends TestBase {

    @Test
    public void testReadIntoPojoList() {
        URL url = getClass().getClassLoader().getResource("countries.csv");
        List<Country> countries = IO.read(url)
                .hint(C.Map("country", "name"))
                .to(new TypeReference<List<Country>>(){});

        eq(1, countries.size());
        Country country = countries.get(0);
        eq(0, country.no);
        eq("AU", country.code);
        eq("Australia", country.name);
    }

    @Test
    public void testReadIntoMapList() {
        URL url = getClass().getClassLoader().getResource("countries.csv");
        List<Map<String, String>> countries = IO.read(url).to(new TypeReference<List<Map<String, String>>>(){});
        eq(1, countries.size());
        Map<String, String> record = countries.get(0);
        eq("0", record.get("no"));
        eq("AU", record.get("code"));
        eq("Australia", record.get("country"));
    }
}
