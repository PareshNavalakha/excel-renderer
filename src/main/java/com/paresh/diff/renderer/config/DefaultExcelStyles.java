package com.paresh.diff.renderer.config;

import org.apache.poi.ss.usermodel.*;

public class DefaultExcelStyles implements ExcelStyles {

    private Workbook workbook;
    private CellStyle titleCellStyle;
    private CellStyle headerCellStyle;
    private CellStyle unchangedCellStyle;
    private CellStyle modifiedCellStyle;
    private CellStyle deletedCellStyle;
    private CellStyle newCellStyle;

    @Override
    public ExcelStyles withWorkBook(Workbook workbook) {
        if (this.workbook == null || this.workbook != workbook) {
            this.workbook = workbook;
            this.titleCellStyle = null;
            this.headerCellStyle = null;
            this.unchangedCellStyle = null;
            this.modifiedCellStyle = null;
            this.deletedCellStyle = null;
            this.newCellStyle = null;
        }
        return this;
    }

    @Override
    public CellStyle getTitleStyle() {
        checkIfWorkbookSet();
        if (titleCellStyle == null) {
            Font titleFont = workbook.createFont();
            titleFont.setFontHeightInPoints((short) 18);
            titleFont.setBold(true);
            titleCellStyle = workbook.createCellStyle();
            titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
            titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            titleCellStyle.setFont(titleFont);
        }
        return titleCellStyle;
    }

    @Override
    public CellStyle getHeaderStyle() {
        checkIfWorkbookSet();
        if (headerCellStyle == null) {
            Font font = workbook.createFont();
            font.setFontHeightInPoints((short) 11);
            font.setColor(IndexedColors.WHITE.getIndex());
            headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
            headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerCellStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerCellStyle.setFont(font);
            headerCellStyle.setWrapText(true);
        }
        return headerCellStyle;
    }

    @Override
    public CellStyle getUnchangedStyle() {
        checkIfWorkbookSet();
        if (unchangedCellStyle == null) {
            unchangedCellStyle = workbook.createCellStyle();
            unchangedCellStyle.setAlignment(HorizontalAlignment.CENTER);
            unchangedCellStyle.setWrapText(true);
            unchangedCellStyle.setBorderRight(BorderStyle.THIN);
            unchangedCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            unchangedCellStyle.setBorderLeft(BorderStyle.THIN);
            unchangedCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            unchangedCellStyle.setBorderTop(BorderStyle.THIN);
            unchangedCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            unchangedCellStyle.setBorderBottom(BorderStyle.THIN);
            unchangedCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        }
        return unchangedCellStyle;
    }

    @Override
    public CellStyle getNewStyle() {
        checkIfWorkbookSet();
        if (newCellStyle == null) {
            newCellStyle = workbook.createCellStyle();
            newCellStyle.setAlignment(HorizontalAlignment.CENTER);
            newCellStyle.setWrapText(true);
            newCellStyle.setBorderRight(BorderStyle.THIN);
            newCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            newCellStyle.setBorderLeft(BorderStyle.THIN);
            newCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            newCellStyle.setBorderTop(BorderStyle.THIN);
            newCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            newCellStyle.setBorderBottom(BorderStyle.THIN);
            newCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            newCellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        return newCellStyle;
    }

    @Override
    public CellStyle getDeletedStyle() {
        checkIfWorkbookSet();
        if (deletedCellStyle == null) {
            deletedCellStyle = workbook.createCellStyle();
            deletedCellStyle.setAlignment(HorizontalAlignment.CENTER);
            deletedCellStyle.setWrapText(true);
            deletedCellStyle.setBorderRight(BorderStyle.THIN);
            deletedCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            deletedCellStyle.setBorderLeft(BorderStyle.THIN);
            deletedCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            deletedCellStyle.setBorderTop(BorderStyle.THIN);
            deletedCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            deletedCellStyle.setBorderBottom(BorderStyle.THIN);
            deletedCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            deletedCellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
            deletedCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        return deletedCellStyle;
    }

    @Override
    public CellStyle getModifiedStyle() {
        checkIfWorkbookSet();
        if (modifiedCellStyle == null) {
            modifiedCellStyle = workbook.createCellStyle();
            modifiedCellStyle.setAlignment(HorizontalAlignment.CENTER);
            modifiedCellStyle.setWrapText(true);
            modifiedCellStyle.setBorderRight(BorderStyle.THIN);
            modifiedCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            modifiedCellStyle.setBorderLeft(BorderStyle.THIN);
            modifiedCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            modifiedCellStyle.setBorderTop(BorderStyle.THIN);
            modifiedCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            modifiedCellStyle.setBorderBottom(BorderStyle.THIN);
            modifiedCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            modifiedCellStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            modifiedCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        return modifiedCellStyle;
    }

    private void checkIfWorkbookSet() {
        if (this.workbook == null) {
            throw new IllegalArgumentException("Workbook should be set using withWorkbook method.");
        }
    }
}
