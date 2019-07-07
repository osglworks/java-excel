package org.osgl.xls;

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Define sheet style
 */
public class SheetStyle {
    /**
     * Should it display grid line on this sheet
     */
    public boolean displayGridLine;

    /**
     * The first row style
     */
    public RowStyle topRowStyle;

    /**
     * Other row style
     */
    public RowStyle dataRowStyle;

    /**
     * An alternative row style - could be used for odd/even interchange along with dataRowStyle
     */
    public RowStyle alternateDataRowStyle;
}
