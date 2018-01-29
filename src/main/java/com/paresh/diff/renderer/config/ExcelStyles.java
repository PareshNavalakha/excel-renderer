package com.paresh.diff.renderer.config;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelStyles {

    ExcelStyles withWorkBook(Workbook workbook);

    CellStyle getTitleStyle();

    CellStyle getHeaderStyle();

    CellStyle getUnchangedStyle();

    CellStyle getNewStyle();

    CellStyle getDeletedStyle();

    CellStyle getModifiedStyle();
}
