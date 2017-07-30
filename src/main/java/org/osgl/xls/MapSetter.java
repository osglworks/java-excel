package org.osgl.xls;

import org.osgl.$;
import org.osgl.Osgl;
import org.osgl.util.E;
import org.osgl.util.PropertySetter;
import org.osgl.util.S;

import java.util.Map;

class MapSetter implements PropertySetter {

    private final String key;

    public MapSetter(String key) {
        E.illegalArgumentIf(S.isBlank(key), "key cannot be blank");
        this.key = key;
    }

    public String key() {
        return key;
    }

    @Override
    public void set(Object entity, Object value, Object index) {
        Map map = $.cast(entity);
        map.put(key, value);
    }

    @Override
    public void setObjectFactory(Osgl.Function<Class<?>, Object> factory) {

    }

    @Override
    public void setStringValueResolver(Osgl.Func2<String, Class<?>, ?> stringValueResolver) {

    }
}
