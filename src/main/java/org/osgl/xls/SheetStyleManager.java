package org.osgl.xls;

/*-
 * #%L
 * Java Excel Tool
 * %%
 * Copyright (C) 2017 - 2019 OSGL (Open Source General Library)
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

import org.apache.poi.hssf.record.cf.BorderFormatting;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.osgl.$;
import org.osgl.logging.LogManager;
import org.osgl.logging.Logger;
import org.osgl.util.E;
import org.osgl.util.IO;
import org.osgl.util.Keyword;
import org.osgl.util.S;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.net.URL;
import java.util.*;

/**
 * Keep {@link SheetStyle} definitions.
 *
 * It can load sheet styles from properties file. A sample sheet style properties file:
 * ```properties
 * sheetStyle.instances=s1,s2
 * sheetStyle.s1.displayGridLine=false
 * sheetStyle.s1.topRow.border=thin,dotted,grey_50_percent
 * sheetStyle.s1.topRow.bgColor=dark_blue
 * sheetStyle.s1.topRow.fgColor=grey_50_percent
 * sheetStyle.s1.dataRow.border=thin,dotted,grey_50_percent
 * sheetStyle.s1.dataRow.bgColor=white
 * sheetStyle.s1.dataRow.fgColor=grey_50_percent
 * sheetStyle.s2.displayGridLine=false
 * sheetStyle.s2.topRow.border=thin,medium_dashed,grey_50_percent
 * sheetStyle.s2.topRow.bgColor=dark_blue
 * sheetStyle.s2.topRow.fgColor=grey_50_percent
 * sheetStyle.s2.dataRow.border=thin,dotted,grey_50_percent
 * sheetStyle.s2.dataRow.bgColor=white
 * sheetStyle.s2.dataRow.fgColor=grey_50_percent
 * ```
 */
public class SheetStyleManager {

    public static final String DEFAULT = "default";

    private static final Logger LOGGER = LogManager.get(SheetStyleManager.class);

    private Map<String, SheetStyle> styleRegistry = new HashMap<>();

    public SheetStyleManager() {
        tryLoadFromResource("sheetstyle.properties", false);
        tryLoadFromResource("sheetstyles.properties", false);
        tryLoadFromResource("sheet_style.properties", false);
        tryLoadFromResource("sheet_styles.properties", false);
        if (!styleRegistry.containsKey(DEFAULT)) {
            tryLoadFromResource("org/osgl/xls/sheet_style.properties", true);
        }
    }

    public void loadFromProperties(Properties properties) {
        String s = properties.getProperty("sheetStyle.instances");
        if (S.isBlank(s)) {
            LOGGER.info("sheetStyle.instances not found, assume only default sheetStyle presented");
            String name = DEFAULT;
            String prefix = null;
            String prefix1 = "sheetStyle.default.";
            String prefix2 = "sheetStyle.";
            for (Object key : properties.keySet()) {
                if (S.string(key).startsWith(prefix1)) {
                    prefix = prefix1;
                    break;
                }
            }
            if (null == prefix) {
                prefix = prefix2;
            }
            SheetStyle style = parse(prefix, properties);
            styleRegistry.put(name, style);
        } else {
            S.List names = S.fastSplit(s, ",");
            for (String name : names) {
                String prefix = "sheetStyle." + name + ".";
                SheetStyle style = parse(prefix, properties);
                styleRegistry.put(name, style);
            }
        }
    }

    public void addSheetStyle(String name, SheetStyle style) {
        styleRegistry.put(name, style);
    }

    public SheetStyle getSheetStyle(String name) {
        SheetStyle sheetStyle = styleRegistry.get(name);
        if (null == sheetStyle) {
            LOGGER.warn("cannot locate sheet style by " + name);
        }
        return sheetStyle;
    }

    public Set<Map.Entry<String, SheetStyle>> allSheetStyles() {
        return styleRegistry.entrySet();
    }

    public void clear() {
        styleRegistry.clear();
    }

    public void loadFromResource(String path) {
        tryLoadFromResource(path, true);
    }

    private void tryLoadFromResource(String path, boolean raiseExceptionIfNotFound) {
        URL url = SheetStyleManager.class.getClassLoader().getResource(path);
        if (null == url) {
            if (raiseExceptionIfNotFound) {
                throw E.unexpected("resource not found: %s", path);
            }
            return;
        }
        Properties properties = IO.loadProperties(url);
        loadFromProperties(properties);
    }

    public static SheetStyle getDefault() {
        return SINGLETON.getSheetStyle(DEFAULT);
    }

    private static Map<String, IndexedColors> COLORS = new HashMap<>(); static {
        for (IndexedColors c : IndexedColors.values()) {
            COLORS.put(c.name(), c);
        }
    }

    private static Map<String, BorderStyle> BORDERS = new HashMap<>(); static {
        for (BorderStyle bs : BorderStyle.values()) {
            BORDERS.put(bs.name(), bs);
        }
    }

    private static IndexedColors colorFor(String name) {
        return IndexedColors.valueOf(name);
    }

    private static BorderStyle borderStyleFor(String name) {
        return BorderStyle.valueOf(name);
    }

    private static RowStyle parseRowStyle(String prefix, Properties properties) {
        RowStyle rowStyle = new RowStyle();
        String s = properties.getProperty(prefix + "border");
        if (S.notBlank(s)) {
            S.List list = S.fastSplit(s, ",");
            for (String trait : list) {
                trait = trait.toUpperCase();
                if (COLORS.containsKey(trait)) {
                    rowStyle.borderColor = colorFor(trait);
                } else if (BORDERS.containsKey(trait)) {
                    rowStyle.borderStyle = borderStyleFor(trait);
                } else {
                    LOGGER.warn("unknown border trait: %s", trait);
                }
            }
        }
        s = properties.getProperty(prefix + "border.top");
        if (S.notBlank(s)) {
            S.List list = S.fastSplit(s, ",");
            for (String trait : list) {
                trait = trait.toUpperCase();
                if (COLORS.containsKey(trait)) {
                    rowStyle.topBorderColor = colorFor(trait);
                } else if (BORDERS.containsKey(trait)) {
                    rowStyle.topBorderStyle = borderStyleFor(trait);
                } else {
                    LOGGER.warn("unknown border trait: %s", trait);
                }
            }
        }
        if (null == rowStyle.topBorderStyle) {
            rowStyle.topBorderStyle = rowStyle.borderStyle;
        }
        if (null == rowStyle.topBorderColor) {
            rowStyle.topBorderColor = rowStyle.borderColor;
        }
        s = properties.getProperty(prefix + "border.left");
        if (S.notBlank(s)) {
            S.List list = S.fastSplit(s, ",");
            for (String trait : list) {
                trait = trait.toUpperCase();
                if (COLORS.containsKey(trait)) {
                    rowStyle.leftBorderColor = colorFor(trait);
                } else if (BORDERS.containsKey(trait)) {
                    rowStyle.leftBorderStyle = borderStyleFor(trait);
                } else {
                    LOGGER.warn("unknown border trait: %s", trait);
                }
            }
        }
        if (null == rowStyle.leftBorderStyle) {
            rowStyle.leftBorderStyle = rowStyle.borderStyle;
        }
        if (null == rowStyle.leftBorderColor) {
            rowStyle.leftBorderColor = rowStyle.borderColor;
        }
        s = properties.getProperty(prefix + "border.right");
        if (S.notBlank(s)) {
            S.List list = S.fastSplit(s, ",");
            for (String trait : list) {
                trait = trait.toUpperCase();
                if (COLORS.containsKey(trait)) {
                    rowStyle.rightBorderColor = colorFor(trait);
                } else if (BORDERS.containsKey(trait)) {
                    rowStyle.rightBorderStyle = borderStyleFor(trait);
                } else {
                    LOGGER.warn("unknown border trait: %s", trait);
                }
            }
        }
        if (null == rowStyle.rightBorderStyle) {
            rowStyle.rightBorderStyle = rowStyle.borderStyle;
        }
        if (null == rowStyle.rightBorderColor) {
            rowStyle.rightBorderColor = rowStyle.borderColor;
        }
        s = properties.getProperty(prefix + "border.bottom");
        if (S.notBlank(s)) {
            S.List list = S.fastSplit(s, ",");
            for (String trait : list) {
                trait = trait.toUpperCase();
                if (COLORS.containsKey(trait)) {
                    rowStyle.bottomBorderColor = colorFor(trait);
                } else if (BORDERS.containsKey(trait)) {
                    rowStyle.bottomBorderStyle = borderStyleFor(trait);
                } else {
                    LOGGER.warn("unknown border trait: %s", trait);
                }
            }
        }
        if (null == rowStyle.bottomBorderStyle) {
            rowStyle.bottomBorderStyle = rowStyle.borderStyle;
        }
        if (null == rowStyle.bottomBorderColor) {
            rowStyle.bottomBorderColor = rowStyle.borderColor;
        }
        s = properties.getProperty(prefix + "bgColor");
        if (S.notBlank(s)) {
            s = s.toUpperCase();
            if (COLORS.containsKey(s)) {
                rowStyle.bgColor = colorFor(s);
            } else {
                LOGGER.warn("unknown bgColor: %s", s);
            }
        }
        s = properties.getProperty(prefix + "fgColor");
        if (S.notBlank(s)) {
            s = s.toUpperCase();
            if (COLORS.containsKey(s)) {
                rowStyle.fgColor = colorFor(s);
            } else {
                LOGGER.warn("unknown fgColor: %s", s);
            }
        }
        return rowStyle.isEmpty() ? null : rowStyle;
    }

    private static SheetStyle parse(String prefix, Properties properties) {
        SheetStyle style = new SheetStyle();
        String s = properties.getProperty(prefix + "displayGridLine");
        if (S.notBlank(s)) {
            style.displayGridLine = Boolean.parseBoolean(s);
        }
        style.topRowStyle = parseRowStyle(prefix + "topRow.", properties);
        style.dataRowStyle = parseRowStyle(prefix + "dataRow.", properties);
        style.alternateDataRowStyle = parseRowStyle(prefix + "alternateDataRow.", properties);
        return style;
    }

    public static final SheetStyleManager SINGLETON = new SheetStyleManager();

}
