package com.paresh.diff.renderer.config;

import com.paresh.diff.renderer.ExcelRendererConstants;
import com.paresh.diff.renderer.RenderingPreferences;
import com.paresh.diff.renderer.util.DefaultExcelStyles;

import java.util.HashMap;
import java.util.Map;

public class ExcelRenderingPreferences implements RenderingPreferences {
    private Map<String, Object> preference = new HashMap<>();

    public void setExcelStyles(ExcelStyles excelStyles) {
        preference.put(ExcelRendererConstants.STYLES, excelStyles);
    }

    public ExcelStyles getExcelStyles() {
        ExcelStyles returnExcelStyle = (ExcelStyles) preference.get(ExcelRendererConstants.STYLES);
        return returnExcelStyle == null ? new DefaultExcelStyles() : returnExcelStyle;
    }

    public String getSheetName() {
        String returnSheetName = (String) preference.get(ExcelRendererConstants.SHEET_NAME);
        return returnSheetName == null ? ExcelRendererConstants.DEFAULT_SHEET_NAME : returnSheetName;
    }

    public void setSheetName(String sheetName) {
        preference.put(ExcelRendererConstants.SHEET_NAME, sheetName);
    }

    public void setTitle(String sheetName) {
        preference.put(ExcelRendererConstants.TITLE, sheetName);
    }

    public String getTitle() {
        return (String) preference.get(ExcelRendererConstants.TITLE);
    }

    @Override
    public void setPreference(String key, Object value) {
        preference.put(key, value);
    }

    @Override
    public Object getPreference(String key) {
        return preference.get(key);
    }

}
