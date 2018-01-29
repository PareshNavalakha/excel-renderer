package com.paresh.diff.renderer.config;

import com.paresh.diff.renderer.ExcelRendererConstants;
import com.paresh.diff.renderer.RenderingPreferences;
import com.paresh.diff.renderer.data.RenderingMode;

import java.util.HashMap;
import java.util.Map;

public class ExcelRenderingPreferences implements RenderingPreferences {
    private Map<String, Object> preference = new HashMap<>();
    private ExcelStyles defaultExcelStyles =  new DefaultExcelStyles();

    public void setExcelStyles(ExcelStyles excelStyles) {
        preference.put(ExcelRendererConstants.STYLES, excelStyles);
    }

    public ExcelStyles getExcelStyles() {
        ExcelStyles returnExcelStyle = (ExcelStyles) preference.get(ExcelRendererConstants.STYLES);
        return returnExcelStyle == null ? defaultExcelStyles : returnExcelStyle;
    }

    public String getSheetName() {
        String returnSheetName = (String) preference.get(ExcelRendererConstants.RECON_SHEET_NAME);
        return returnSheetName == null ? RenderingDefaults.getSheetName(getRenderingMode()) : returnSheetName;
    }

    public void setSheetName(String sheetName) {
        preference.put(ExcelRendererConstants.RECON_SHEET_NAME, sheetName);
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

    public RenderingMode getRenderingMode() {
        RenderingMode renderingMode = (RenderingMode) preference.get(ExcelRendererConstants.RENDERING_MODE);
        return renderingMode == null ? RenderingMode.DIFF : renderingMode;
    }

    public void setRenderingMode(RenderingMode renderingMode) {
        preference.put(ExcelRendererConstants.RENDERING_MODE, renderingMode);
    }

    public String getSource1() {
        String source = (String) preference.get(ExcelRendererConstants.SOURCE1_NAME);
        return source == null ? ExcelRendererConstants.DEFAULT_SOURCE1_NAME : source;
    }

    public void setSource1(String source) {
        preference.put(ExcelRendererConstants.SOURCE1_NAME, source);
    }

    public String getSource2() {
        String source = (String) preference.get(ExcelRendererConstants.SOURCE2_NAME);
        return source == null ? ExcelRendererConstants.DEFAULT_SOURCE2_NAME : source;
    }

    public void setSource2(String source) {
        preference.put(ExcelRendererConstants.SOURCE2_NAME, source);
    }

    public boolean isChangeTypeHeaderRequired() {
        Boolean changeTypeRequired = (Boolean) preference.get(ExcelRendererConstants.CHANGE_TYPE_REQUIRED);
        return changeTypeRequired == null ? true : changeTypeRequired;
    }

    public void changeTypeHeaderRequired(boolean changeTypeRequired) {
        preference.put(ExcelRendererConstants.CHANGE_TYPE_REQUIRED, changeTypeRequired);
    }

    public boolean isSummaryOfChangeHeaderRequired() {
        Boolean summaryOfChangeHeaderRequired = (Boolean) preference.get(ExcelRendererConstants.SUMMARY_OF_CHANGE_REQUIRED);
        return summaryOfChangeHeaderRequired == null ? true : summaryOfChangeHeaderRequired;
    }

    public void summaryOfChangeHeaderRequired(boolean summaryOfChangeHeaderRequired) {
        preference.put(ExcelRendererConstants.SUMMARY_OF_CHANGE_REQUIRED, summaryOfChangeHeaderRequired);
    }


    public String getBeforeSheetName() {
        return RenderingDefaults.getBeforeSheetName(getRenderingMode(),getSource1());
    }

    public String getAfterSheetName() {
        return RenderingDefaults.getAfterSheetName(getRenderingMode(),getSource2());
    }
}
