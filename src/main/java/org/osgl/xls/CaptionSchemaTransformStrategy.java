package org.osgl.xls;

import org.osgl.$;
import org.osgl.Osgl;
import org.osgl.util.Keyword;

import java.util.Map;

public final class CaptionSchemaTransformStrategy {
    private CaptionSchemaTransformStrategy() {}

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
