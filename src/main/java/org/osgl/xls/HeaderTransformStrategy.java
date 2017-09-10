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

import org.osgl.$;
import org.osgl.Osgl;
import org.osgl.util.Keyword;

import java.util.Map;

public final class HeaderTransformStrategy {
    private HeaderTransformStrategy() {}

    public static $.Function<String, String> AS_CAPTION = $.F.identity();

    public static $.Function<String, String> TO_JAVA_NAME = transformTo(Keyword.Style.JAVA_VARIABLE);

    public static $.Function<String, String> transformTo(final Keyword.Style style) {
        return new $.Transformer<String, String>() {
            @Override
            public String transform(String s) {
                return style.toString(Keyword.of(s));
            }
        };
    }

    public static $.Function<String, String> translate(final Map<String, String> dictionary) {
        return new Osgl.Transformer<String, String>() {
            @Override
            public String transform(String caption) {
                return dictionary.get(caption);
            }
        };
    }
}
