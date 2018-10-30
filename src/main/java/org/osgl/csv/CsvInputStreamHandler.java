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

import static org.osgl.csv.CsvParser.parse;

import org.osgl.$;
import org.osgl.util.*;

import java.io.InputStream;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;
import javax.inject.Singleton;

@Singleton
public class CsvInputStreamHandler implements IO.InputStreamHandler {

    private static Type DIRECT_TYPE_1 = new TypeReference<List<Map<String, String>>>(){}.getType();
    private static Type DIRECT_TYPE_2 = new TypeReference<List<LinkedHashMap<String, String>>>(){}.getType();
    private static Type DIRECT_TYPE_3 = new TypeReference<List<Map<String, Object>>>(){}.getType();
    private static Type DIRECT_TYPE_4 = new TypeReference<List<LinkedHashMap<String, Object>>>(){}.getType();
    private static Type DIRECT_TYPE_5 = new TypeReference<List<Map>>(){}.getType();
    private static Type DIRECT_TYPE_6 = new TypeReference<List<LinkedHashMap>>(){}.getType();

    private static IdentityHashMap<Type, Type> directTypeLookup = new IdentityHashMap<>(C.<Type, Type>Map(
            DIRECT_TYPE_1, DIRECT_TYPE_1,
            DIRECT_TYPE_2, DIRECT_TYPE_2,
            DIRECT_TYPE_3, DIRECT_TYPE_3,
            DIRECT_TYPE_4, DIRECT_TYPE_4,
            DIRECT_TYPE_5, DIRECT_TYPE_5,
            DIRECT_TYPE_6, DIRECT_TYPE_6));

    @Override
    public boolean support(Type type) {
        $.Pair<Class, Class> containerElement = parseType(type);
        Class container = containerElement.left();
        Class element = containerElement.right();
        return null != container && !$.isSimpleType(element);
    }

    @Override
    public <T> T read(InputStream inputStream, Type type, MimeType mimeType, Object hint) {
        List<String> lines = IO.readLines(inputStream);
        List<Map<String, String>> temp = parse(lines);
        if (directTypeLookup.containsKey(type)) {
            return (T) temp;
        }
        $.Pair<Class, Class> containerElement = parseType(type);
        $._MappingStage stage = $.map(temp).targetGenericType(type);
        if (hint instanceof Map) {
            stage.withHeadMapping((Map) hint);
        }
        return (T) stage.to(containerElement.first());
    }

    $.Pair<Class, Class> parseType(Type type) {
        if (type instanceof Class) {
            Class clz = (Class) type;
            if (Collection.class.isAssignableFrom(clz)) {
                return $.Pair(clz, (Class) Object.class);
            }
        } else if (type instanceof ParameterizedType) {
            ParameterizedType ptype = (ParameterizedType) type;
            Type rawType = ptype.getRawType();
            if (rawType instanceof Class) {
                Class container = (Class) rawType;
                if (Collection.class.isAssignableFrom(container)) {
                    Type[] typeParams = ptype.getActualTypeArguments();
                    if (null != typeParams && typeParams.length == 1) {
                        Type elementType = typeParams[0];
                        if (elementType instanceof Class) {
                            return $.Pair(container, (Class) elementType);
                        } else if (elementType instanceof ParameterizedType) {
                            ParameterizedType pElementType = (ParameterizedType) elementType;
                            Type rawElementType = pElementType.getRawType();
                            if (rawElementType instanceof Class) {
                                return $.Pair(container, (Class) rawElementType);
                            }
                        }
                    }
                }
            }
        }
        return $.Pair(null, null);
    }
}
