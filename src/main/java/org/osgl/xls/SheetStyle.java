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
