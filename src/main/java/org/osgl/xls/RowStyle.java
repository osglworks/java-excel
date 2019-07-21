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

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.File;

public class RowStyle {
    public BorderStyle borderStyle;
    public IndexedColors borderColor;

    public BorderStyle topBorderStyle;
    public BorderStyle rightBorderStyle;
    public BorderStyle bottomBorderStyle;
    public BorderStyle leftBorderStyle;

    public IndexedColors topBorderColor;
    public IndexedColors rightBorderColor;
    public IndexedColors bottomBorderColor;
    public IndexedColors leftBorderColor;

    public IndexedColors bgColor;
    public IndexedColors fgColor;

    public boolean isEmpty() {
        return borderStyle == null && null == borderColor && null == bgColor && null == fgColor;
    }


    public static void main(String[] args) {
        File file = new File("/tmp/1/1.xlsx");
        new ExcelReader.Builder().file(file).build().read();
    }


}
