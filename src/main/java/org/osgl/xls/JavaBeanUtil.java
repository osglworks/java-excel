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
import org.osgl.util.C;
import org.osgl.util.PropertySetter;
import org.osgl.util.S;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

class JavaBeanUtil {
    private JavaBeanUtil() {}

    static Map<String, PropertySetter> setters(Class<?> schema, Map<String, String> headerMapping) {
        Map<String, PropertySetter> setterMap = new HashMap<>();
        loadSettersFromSetterMethod(schema, setterMap);
        loadSettersFromFields(schema, setterMap);
        processNestedSetters(schema, headerMapping, setterMap);
        return setterMap;
    }

    private static void processNestedSetters(Class<?> schema, Map<String, String> headerMapping, Map<String, PropertySetter> setterMap) {
        Set<String> nestedProperties = C.newSet();
        for (String prop: headerMapping.values()) {
            if (prop.contains(".")) {
                nestedProperties.add(prop);
            }
        }
        if (nestedProperties.isEmpty()) {
            return;
        }
        for (final String prop : nestedProperties) {
            PropertySetter setter = new PropertySetter() {
                @Override
                public void set(Object entity, Object value, Object index) {
                    $.setProperty(entity, value, prop);
                }

                @Override
                public void setObjectFactory(Osgl.Function<Class<?>, Object> factory) {
                }

                @Override
                public void setStringValueResolver(Osgl.Func2<String, Class<?>, ?> stringValueResolver) {
                }
            };
            setterMap.put(prop, setter);
        }
    }

    private static void loadSettersFromSetterMethod(Class<?> schema, Map<String, PropertySetter> map) {
        Method[] methods = schema.getMethods();
        for (Method method : methods) {
            String name = method.getName();
            if (name.startsWith("set") && !Void.class.equals(method.getReturnType())) {
                String property = name.substring(3);
                if (property.length() < 1) {
                    continue;
                }
                char c = property.charAt(0);
                if (c > 'Z' || c < 'A') {
                    continue;
                }
                String propertyName = S.lowerFirst(property);
                map.put(propertyName, $.propertyHandlerFactory.createPropertySetter(schema, propertyName));
            }
        }
    }

    private static void loadSettersFromFields(Class<?> schema, Map<String, PropertySetter> map) {
        Field[] fields = schema.getDeclaredFields();
        for (Field field : fields) {
            int modifier = field.getModifiers();
            if (!Modifier.isPublic(modifier) || Modifier.isStatic(modifier) || Modifier.isTransient(modifier)) {
                continue;
            }
            String propertyName = field.getName();
            map.put(propertyName, $.propertyHandlerFactory.createPropertySetter(schema, propertyName));
        }
    }
}
