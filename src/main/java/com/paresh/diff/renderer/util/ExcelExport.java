package com.paresh.diff.renderer.util;

import com.paresh.diff.cache.ClassMetadataCache;
import com.paresh.diff.dto.ChangeType;
import com.paresh.diff.dto.ClassMetadata;
import com.paresh.diff.dto.Diff;
import com.paresh.diff.dto.DiffResponse;
import com.paresh.diff.renderer.ExcelRendererConstants;
import com.paresh.diff.renderer.config.ExcelRenderingPreferences;
import com.paresh.diff.util.ReflectionUtil;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.Collection;
import java.util.List;

public class ExcelExport {

    private static final Logger logger = LoggerFactory.getLogger(ExcelExport.class);

    public static void exportToExcel(String fileName, Object before, Object after, DiffResponse diffResponse, ExcelRenderingPreferences renderingPreferences) {

        List consolidatedList = RenderingUtil.getConsolidatedCollection(before, after, diffResponse);

        if (!CollectionUtils.isEmpty(consolidatedList)) {
            int rowNumber = 0;
            Class collectionElementClass = ReflectionUtil.getCollectionElementClass(before, after);
            ClassMetadata classMetadata = diffResponse.getClassMetaDataMap().get(collectionElementClass);

            List<String> headers = RenderingUtil.getHeaders(classMetadata);

            try (Workbook workbook = new XSSFWorkbook();) {

                Sheet sheet = createSheet(workbook, renderingPreferences);

                generateTitleRow(workbook, sheet, RenderingUtil.getTitle(renderingPreferences, collectionElementClass, classMetadata), renderingPreferences, headers.size()-1);
                generateHeaderRows(workbook, sheet, headers, ++rowNumber, renderingPreferences);

                ChangeType changeType;
                String changeTypeText = null;

                int outerIndex = 0;

                for (Diff diff : diffResponse.getDiffs()) {
                    changeType = diff.getChangeType();
                    switch (changeType) {
                        case ADDED:
                            changeTypeText = "New";
                            break;
                        case UPDATED:
                            changeTypeText = "Updated";
                            break;
                        case DELETED:
                            changeTypeText = "Deleted";
                            break;
                        case NO_CHANGE:
                            changeTypeText = ExcelRendererConstants.BLANK;
                            break;
                    }
                    generateCellContent(sheet, ++rowNumber, 0, changeTypeText, ExcelRendererConstants.BLANK, getCellStyle(workbook,changeType, renderingPreferences));

                    List<Method> methods = classMetadata.getMethods();
                    ChangeType attributeChangeType;
                    Diff relevantDiff;
                    for (int innerIndex = 0; innerIndex < methods.size(); innerIndex++) {
                        relevantDiff = RenderingUtil.getRelevantDiff(diff.getChildDiffs(), methods.get(innerIndex), classMetadata);
                        if (relevantDiff != null) {
                            attributeChangeType = relevantDiff.getChangeType();
                        } else {
                            attributeChangeType = ChangeType.NO_CHANGE;
                        }
                        generateCellContent(sheet, rowNumber, innerIndex + 1, RenderingUtil.getAttributeObject(methods.get(innerIndex), consolidatedList.get(outerIndex)), getComment(diff, methods.get(innerIndex), attributeChangeType, consolidatedList.get(outerIndex), before, after, classMetadata), getCellStyle(workbook, attributeChangeType, renderingPreferences));
                    }
                    outerIndex++;
                }
                FileOutputStream out = new FileOutputStream(fileName);
                workbook.write(out);
                out.close();
                logger.info("Done saving file {}", fileName);
            } catch (IOException e) {
                logger.error("Exception while exporting excel", e);
            }
        }
    }

    private static String getComment(Diff classDiff, Method method, ChangeType changeType, Object object, Object before, Object after, ClassMetadata classMetadata) {
        String response = null;

        Collection beforeCollection = (Collection) before;
        Collection afterCollection = (Collection) after;

        Object afterObject = null;
        Object beforeObject = null;

        switch (changeType) {
            case UPDATED:
                beforeObject = ClassMetadataCache.getInstance().getObjectFromIdentifier(classDiff.getIdentifier(), beforeCollection);
                response = "Original value: " + ReflectionUtil.getMethodResponse(method, beforeObject);
                break;
            case DELETED:
                afterObject = ClassMetadataCache.getInstance().getObjectFromIdentifier(classDiff.getIdentifier(), afterCollection);
                response = "Original value: " + ReflectionUtil.getMethodResponse(method, afterObject);
                break;
            case NO_CHANGE:
            case ADDED:
            default:
                break;
        }
        return response;
    }


    private static Sheet createSheet(Workbook workbook, ExcelRenderingPreferences renderingPreferences) {
        return workbook.createSheet(renderingPreferences.getSheetName());
    }

    private static void generateTitleRow(Workbook workbook, Sheet sheet, String title, ExcelRenderingPreferences renderingPreferences, int headerCount) {
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(renderingPreferences.getExcelStyles().getTitleStyle(workbook));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headerCount));
    }

    private static void generateHeaderRows(Workbook workbook, Sheet sheet, List<String> headers, int row, ExcelRenderingPreferences
            renderingPreferences) {
        Row headerRow = sheet.createRow(row);
        headerRow.setHeightInPoints(40);
        Cell headerCell;
        for (int index = 0; index < headers.size(); index++) {
            headerCell = headerRow.createCell(index);
            headerCell.setCellValue(headers.get(index));
            headerCell.setCellStyle(renderingPreferences.getExcelStyles().getHeaderStyle(workbook));
        }
    }

    private static CellStyle getCellStyle(Workbook workbook, ChangeType changeType, ExcelRenderingPreferences
            renderingPreferences) {
        switch (changeType) {
            case ADDED:
                return renderingPreferences.getExcelStyles().getNewStyle(workbook);
            case UPDATED:
                return renderingPreferences.getExcelStyles().getModifiedStyle(workbook);
            case DELETED:
                return renderingPreferences.getExcelStyles().getDeletedStyle(workbook);
            case NO_CHANGE:
                return renderingPreferences.getExcelStyles().getUnchangedStyle(workbook);
            default:
                return renderingPreferences.getExcelStyles().getUnchangedStyle(workbook);
        }
    }

    private static void generateCellContent(Sheet sheet, int row, int column, Object cellValue, String comment, CellStyle cellStyle) {
        Row cellRow = sheet.getRow(row);
        if (cellRow == null) {
            cellRow = sheet.createRow(row);
        }
        Cell cell = cellRow.createCell(column);
        //TODO improve
        if (cellValue != null) {
            cell.setCellValue(cellValue.toString());

        }
        if (comment != null && !comment.isEmpty()) {
            cell.setCellComment(generateCellComment((XSSFSheet) sheet, comment, column, row));
        }
        cell.setCellStyle(cellStyle);
    }

    private static Comment generateCellComment(XSSFSheet sheet, String text, int column, int row) {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, column, row, column + 1, row + 1);
        XSSFComment comment = drawing.createCellComment(anchor);

        // set text in the comment
        comment.setString(new XSSFRichTextString(text));

        //set comment author.
        //you can see it in the status bar when moving mouse over the commented cell
        comment.setAuthor("Recon Util");

        return comment;
    }
}
