package org.osgl.xls;

/*-
 * #%L
 * Java Excel Reader
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

import org.osgl.$;
import org.osgl.util.*;

import java.io.InputStream;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.List;
import java.util.Map;

public class ExcelInputStreamHandler implements IO.InputStreamHandler {

    @Override
    public boolean support(Type type) {
        if (type instanceof Class) {
            Class clazz = (Class) type;
            return Map.class.isAssignableFrom(clazz) || List.class.isAssignableFrom(clazz);
        } else if (type instanceof ParameterizedType){
            ParameterizedType ptype = $.cast(type);
            Type rawType = ptype.getRawType();
            return support(rawType);
        }
        return false;
    }

    @Override
    public <T> T read(InputStream is, Type targetType, MimeType mimeType, Object hint) {
        if (targetType instanceof Class) {
            Class clazz = (Class) targetType;
            if (Map.class.isAssignableFrom(clazz)) {
                return (T) readIntoMap(is, Map.class, mimeType, hint);
            } else if (List.class.isAssignableFrom(clazz)) {
                return (T) readIntoList(is, List.class, mimeType, hint);
            } else {
                throw new UnsupportedOperationException();
            }
        } else if (targetType instanceof ParameterizedType){
            E.unsupportedIfNot(support(targetType));
            ParameterizedType ptype = $.cast(targetType);
            Type rawType = ptype.getRawType();
            boolean isMap = Map.class.isAssignableFrom((Class) rawType);
            if (isMap) {
                Type[] typeParams = ptype.getActualTypeArguments();
                E.unsupportedIf(typeParams.length != 2);
                Type keyType = typeParams[0];
                E.unsupportedIf(keyType != String.class);
                return (T) readIntoMap(is, targetType, mimeType, hint);
            } else {
                return (T) readIntoList(is, targetType, mimeType, hint);
            }
        } else {
            throw E.unsupport();
        }
    }

    private Map readIntoMap(InputStream is, Type targetType, MimeType mimeType, Object hint) {
        return readIntoMap(reader(is, mimeType, hint), targetType);
    }

    private Map readIntoMap(ExcelReader reader, Type type) {
        if (type instanceof ParameterizedType) {
            ParameterizedType ptype = $.cast(type);
            Type[] typeParams = ptype.getActualTypeArguments();
            Type valType = typeParams[1];
            if (valType instanceof Class) {
                return reader.readSheets((Class) valType);
            }
        }
        return reader.readSheets();
    }

    private List readIntoList(InputStream is, Type targetType, MimeType mimeType, Object hint) {
        return readIntoList(reader(is, mimeType, hint), targetType);
    }

    private List readIntoList(ExcelReader reader, Type type) {
        if (type instanceof ParameterizedType) {
            ParameterizedType ptype = $.cast(type);
            Type[] typeParams = ptype.getActualTypeArguments();
            Type valType = typeParams[0];
            if (valType instanceof Class) {
                return reader.read((Class) valType);
            }
        }
        return reader.read();
    }

    private ExcelReader reader(InputStream is, MimeType mimeType, Object hint) {
        ExcelReader.Builder builder;
        if (hint instanceof ExcelReader.Builder) {
            builder = $.cast(hint);
            builder.mimeType(mimeType);
            builder.inputStream(is);
        } else {
            builder = ExcelReader.builder().mimeType(mimeType).inputStream(is);
        }
        return builder.build();
    }
}
