package org.osgl.xls;

import org.osgl.$;
import org.osgl.util.PropertySetter;
import org.osgl.util.S;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.HashMap;
import java.util.Map;

class JavaBeanUtil {
    private JavaBeanUtil() {}

    static Map<String, PropertySetter> setters(Class<?> schema) {
        Map<String, PropertySetter> map = new HashMap<>();
        loadSettersFromSetterMethod(schema, map);
        loadSettersFromFields(schema, map);
        return map;
    }

    private static void loadSettersFromSetterMethod(Class<?> schema, Map<String, PropertySetter> map) {
        Method[] methods = schema.getMethods();
        for (Method method : methods) {
            String name = method.getName();
            if ("getClass".equals(name)) {
                continue;
            }
            if (name.startsWith("get") && !Void.class.equals(method.getReturnType())) {
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
