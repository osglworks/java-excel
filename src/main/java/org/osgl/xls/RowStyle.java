package org.osgl.xls;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

public class RowStyle {
    public BorderStyle borderStyle;
    public IndexedColors borderColor;
    public IndexedColors bgColor;
    public IndexedColors fgColor;

    public boolean isEmpty() {
        return borderStyle == null && null == borderColor && null == bgColor && null == fgColor;
    }

}
