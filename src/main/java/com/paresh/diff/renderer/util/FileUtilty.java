package com.paresh.diff.renderer.util;

import java.io.File;

public class FileUtilty {

    public static String getTemporaryLocation() {
        return System.getProperty("java.io.tmpdir");
    }

    public static String getSeperator() {
        return File.separator;
    }

}
