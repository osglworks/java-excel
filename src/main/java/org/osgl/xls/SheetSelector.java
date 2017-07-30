package org.osgl.xls;

import org.apache.poi.ss.usermodel.Sheet;
import org.osgl.$;
import org.osgl.util.E;
import org.osgl.util.S;

import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Utility to define sheet select logic
 */
public class SheetSelector {

    /**
     * Select all sheets
     */
    public static $.Predicate<Sheet> ALL = new $.Predicate<Sheet>() {
        @Override
        public boolean test(Sheet rows) {
            return true;
        }
    };

    /**
     * Select sheets by name list. Name list is case insensitive
     *
     * The name list could be provided using string array or a single string
     * with name list separated by comma `,`
     *
     * @param names the name list
     * @return a predicate to select sheets
     */
    public static $.Predicate<Sheet> byName(final String ... names) {
        E.illegalArgumentIf(names.length == 0, "name list expected");
        final Set<String> nameList = new HashSet<>();
        for (String s : names) {
            if (S.blank(s)) {
                continue;
            }
            s = s.trim().toLowerCase();
            List<String> splited = S.fastSplit(s, ",");
            nameList.addAll(splited);
        }
        return new $.Predicate<Sheet>() {
            @Override
            public boolean test(Sheet sheet) {
                return nameList.contains(sheet.getSheetName().toLowerCase());
            }
        };
    }

    /**
     * Select sheets by exclude name list. Name list is case insensitive
     *
     * The name list could be provided using string array or a single string
     * with name list separated by comma `,`
     *
     * @param names the exclude name list
     * @return a predicate to select sheets
     */
    public static $.Predicate<Sheet> excludeByName(final String... names) {
        return byName(names).negate();
    }

    /**
     * Select sheets by index list. Note index number is `0` based
     *
     * @param indexes the index list
     * @return a predicate to select sheets
     */
    public static $.Predicate<Sheet> byPosition(final int ... indexes) {
        E.illegalArgumentIf(indexes.length == 0, "index list expected");
        Arrays.sort(indexes);
        return new $.Predicate<Sheet>() {
            @Override
            public boolean test(Sheet sheet) {
                int myIndex = sheet.getWorkbook().getSheetIndex(sheet);
                return Arrays.binarySearch(indexes, myIndex) > - 1;
            }
        };
    }

    /**
     * Select sheets by exclude index list. Note index number is `0` based
     *
     * @param indexes the exclude index list
     * @return a predicate to select sheets
     */
    public static $.Predicate<Sheet> excludeByPosition(final int... indexes) {
        return byPosition(indexes).negate();
    }

}
