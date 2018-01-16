package com.paresh.diff.renderer.util;

import com.paresh.diff.dto.DiffResponse;
import com.paresh.diff.renderer.ExcelRendererConstants;
import com.paresh.diff.renderer.config.ExcelRenderingPreferences;
import com.paresh.diff.util.ReflectionUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public class ExcelExport {

    public static void exportToExcel(String fileName, Object before, Object after, DiffResponse diffResponse, ExcelRenderingPreferences renderingPreferences) {


    }

    private static String getTitle(ExcelRenderingPreferences renderingPreferences, Object before, Object after, DiffResponse diffResponse) {
        String title = renderingPreferences.getTitle();
        if (title == null) {
            Class clazz = ReflectionUtil.getCollectionElementClass(before);
            if (clazz == null) {
                clazz = ReflectionUtil.getCollectionElementClass(after);
            }
            title = diffResponse.getClassMetaDataMap().get(clazz).getClassDescription();
        }
        return title;
    }

    private static Sheet createSheet(Workbook workbook, ExcelRenderingPreferences renderingPreferences) {
        return workbook.createSheet(renderingPreferences.getSheetName());
    }

    private static Workbook generateWorkbook() {
        return new XSSFWorkbook();
    }

    private static void generateTitleRows(Sheet sheet, String title, ExcelRenderingPreferences renderingPreferences) {
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.TITLE_STYLE));
        sheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$L$1"));
    }

    private static void generateHeaderRows(Sheet sheet, List<String> headers, int row, ExcelRenderingPreferences renderingPreferences) {
        Row headerRow = sheet.createRow(row);
        headerRow.setHeightInPoints(40);
        Cell headerCell;
        for (int index = 0; index < headers.size(); index++) {
            headerCell = headerRow.createCell(index);
            headerCell.setCellValue(headers.get(index));
            headerCell.setCellStyle(renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.HEADER_STYLE));
        }
    }


}
