package com.paresh.diff.renderer.config;

import com.paresh.diff.renderer.data.RenderingMode;

public class RenderingDefaults {

    public static String getNewDescription(RenderingMode mode, String source) {
        if (mode != null && mode.equals(RenderingMode.RECON)) {
            return "Present only in " + source;
        } else {
            return "New";
        }
    }

    public static String getDeletedDescription(RenderingMode mode, String source) {
        if (mode != null && mode.equals(RenderingMode.RECON)) {
            return "Absent in " + source;
        } else {
            return "Deleted";
        }
    }

    public static String getUpdatedDescription(RenderingMode mode) {

        if (mode != null && mode.equals(RenderingMode.RECON)) {
            return "Different";
        } else {
            return "Modified";
        }
    }

    public static String getSheetName(RenderingMode mode) {

        if (mode != null && mode.equals(RenderingMode.RECON)) {
            return "Reconciliation Report";
        } else {
            return "Difference Report";
        }
    }

    public static String getTitlePostfix(RenderingMode mode) {

        if (mode != null && mode.equals(RenderingMode.RECON)) {
            return " Reconciliation";
        } else {
            return " Difference";
        }
    }

    public static String getBeforeSheetName(RenderingMode renderingMode, String source1) {
        if(renderingMode.equals(RenderingMode.DIFF))
        {
            return "Before";
        }
        else
        {
            return source1;
        }
    }

    public static String getAfterSheetName(RenderingMode renderingMode, String source2) {
        if(renderingMode.equals(RenderingMode.DIFF))
        {
            return "After";
        }
        else
        {
            return source2;
        }
    }
}
