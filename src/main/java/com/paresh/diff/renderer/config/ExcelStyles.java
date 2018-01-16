package com.paresh.diff.renderer.config;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelStyles {
    public void createStyles(Workbook workbook);
    public CellStyle getStyle(String styleName);
}
