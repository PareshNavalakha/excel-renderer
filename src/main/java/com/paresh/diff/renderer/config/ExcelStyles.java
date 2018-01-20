package com.paresh.diff.renderer.config;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelStyles {

    CellStyle getTitleStyle(Workbook workbook);

    CellStyle getHeaderStyle(Workbook workbook);

    CellStyle getUnchangedStyle(Workbook workbook);

    CellStyle getNewStyle(Workbook workbook);

    CellStyle getDeletedStyle(Workbook workbook);

    CellStyle getModifiedStyle(Workbook workbook);
}
