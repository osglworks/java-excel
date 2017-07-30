package org.osgl.xls;

class Feature {

    private static Boolean supportXSSF;

    static {
        try {
            Class.forName("org.apache.poi.xssf.usermodel.XSSFWorkbook");
            supportXSSF = true;
        } catch (ClassNotFoundException e) {
            supportXSSF = false;
        }
    }

    static boolean isSupportXSSF() {
        return supportXSSF;
    }

    static void checkXlsxSupport() {
        if (!supportXSSF) {
            throw new UnsupportedOperationException("Office 2007 file format not support. Make sure your project has added poi-ooxml library");
        }
    }

    static void checkSupport(boolean isXlsx) {
        if (isXlsx) {
            checkXlsxSupport();
        }
    }
}
