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
    public void setObjectFactory($.Function<Class<?>, Object> factory) {

    }

    @Override
    public void setStringValueResolver($.Func2<String, Class<?>, ?> stringValueResolver) {

    }
}
