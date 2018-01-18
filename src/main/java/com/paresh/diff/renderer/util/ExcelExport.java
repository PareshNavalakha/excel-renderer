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
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class ExcelExport {

    private static final Logger logger = LoggerFactory.getLogger(ExcelExport.class);

    public static void exportToExcel(String fileName, Object before, Object after, DiffResponse diffResponse, ExcelRenderingPreferences renderingPreferences) {

        List consolidatedList = getConsolidatedCollection(before, after, diffResponse);

        if (!CollectionUtils.isEmpty(consolidatedList)) {
            int rowNumber = 0;
            Class collectionElementClass = ReflectionUtil.getCollectionElementClass(before, after);
            ClassMetadata classMetadata = diffResponse.getClassMetaDataMap().get(collectionElementClass);

            List<String> headers = getHeaders(classMetadata);

            try (Workbook workbook = new XSSFWorkbook();) {

                Sheet sheet = createSheet(workbook, renderingPreferences);

                generateTitleRow(sheet, getTitle(renderingPreferences, collectionElementClass, classMetadata), renderingPreferences);
                generateHeaderRows(sheet, headers, ++rowNumber, renderingPreferences);

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
                    generateCellContent(sheet, ++rowNumber, 0, changeTypeText, ExcelRendererConstants.BLANK, getCellStyle(changeType, renderingPreferences));

                    List<Method> methods = classMetadata.getMethods();
                    for (int innerIndex = 0; innerIndex < methods.size(); innerIndex++) {
                        generateCellContent(sheet, rowNumber, innerIndex + 1, getAttributeObject(methods.get(innerIndex), consolidatedList.get(outerIndex)), getComment(diff, methods.get(innerIndex), diff.getChildDiffs(), consolidatedList.get(outerIndex), before, after, classMetadata), getCellStyle(diff.getChangeType(), renderingPreferences));
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

    private static Object getAttributeObject(Method method, Object o) {
        return ReflectionUtil.getMethodResponse(method, o);
    }

    private static String getComment(Diff classDiff, Method method, Collection<Diff> childDiffs, Object object, Object before, Object after, ClassMetadata classMetadata) {
        String response = null;

        if (!CollectionUtils.isEmpty(childDiffs)) {
            Diff methodDiff = getRelevantDiff(childDiffs, method, classMetadata);
            if (methodDiff != null) {
                Collection beforeCollection = (Collection) before;
                Collection afterCollection = (Collection) after;

                Object afterObject = null;
                Object beforeObject = null;

                switch (methodDiff.getChangeType()) {
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
                if (response != null) {
                    return "Original value: " + response;
                }
            }
        }
        return response;
    }

    private static Diff getRelevantDiff(Collection<Diff> childDiffs, Method method, ClassMetadata classMetadata) {
        String description = classMetadata.getMethodDescriptions().get(classMetadata.getMethods().indexOf(method));
        for (Diff diff : childDiffs) {
            if (description.equals(diff.getFieldDescription())) {
                return diff;
            }
        }
        return null;
    }


    //TODO Implement ordering
    private static List<String> getHeaders(ClassMetadata classMetadata) {
        List<String> headers = null;
        if (!CollectionUtils.isEmpty(classMetadata.getMethodDescriptions())) {
            headers = new ArrayList<>(classMetadata.getMethodDescriptions().size() + 1);
            headers.add(ExcelRendererConstants.DIFF_STATUS_HEADER);
            headers.addAll(classMetadata.getMethodDescriptions());
        }
        return headers;
    }

    private static List getConsolidatedCollection(Object before, Object after, DiffResponse diffResponse) {
        List consolidatedCollection = new ArrayList(diffResponse.getDiffs().size());
        Collection beforeCollection = (Collection) before;
        Collection afterCollection = (Collection) after;

        for (Diff diff : diffResponse.getDiffs()) {
            switch (diff.getChangeType()) {
                case ADDED:
                case UPDATED:
                    consolidatedCollection.add(ClassMetadataCache.getInstance().getObjectFromIdentifier(diff.getIdentifier(), afterCollection));
                    break;
                case DELETED:
                case NO_CHANGE:
                    consolidatedCollection.add(ClassMetadataCache.getInstance().getObjectFromIdentifier(diff.getIdentifier(), beforeCollection));
                    break;
            }
        }
        return consolidatedCollection;
    }

    private static String getTitle(ExcelRenderingPreferences renderingPreferences, Class collectionElementClass, ClassMetadata classMetadata) {
        String title = renderingPreferences.getTitle();
        if (title == null) {
            title = classMetadata.getClassDescription();
        }
        return title;
    }

    private static Sheet createSheet(Workbook workbook, ExcelRenderingPreferences renderingPreferences) {
        return workbook.createSheet(renderingPreferences.getSheetName());
    }

    private static void generateTitleRow(Sheet sheet, String title, ExcelRenderingPreferences renderingPreferences) {
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.TITLE_STYLE));
        sheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$L$1"));
    }

    private static void generateHeaderRows(Sheet sheet, List<String> headers, int row, ExcelRenderingPreferences
            renderingPreferences) {
        Row headerRow = sheet.createRow(row);
        headerRow.setHeightInPoints(40);
        Cell headerCell;
        for (int index = 0; index < headers.size(); index++) {
            headerCell = headerRow.createCell(index);
            headerCell.setCellValue(headers.get(index));
            headerCell.setCellStyle(renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.HEADER_STYLE));
        }
    }

    private static CellStyle getCellStyle(ChangeType changeType, ExcelRenderingPreferences
            renderingPreferences) {
        switch (changeType) {
            case ADDED:
                return renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.NEW_CELL_STYLE);
            case UPDATED:
                return renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.MODIFY_CELL_STYLE);
            case DELETED:
                return renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.DELETE_CELL_STYLE);
            case NO_CHANGE:
                return renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.NORMAL_CELL_STYLE);
            default:
                return renderingPreferences.getExcelStyles().getStyle(ExcelRendererConstants.NORMAL_CELL_STYLE);
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
            cell.setCellComment(getComment((XSSFSheet) sheet, comment, column, row));
        }
        cell.setCellStyle(cellStyle);
    }

    private static Comment getComment(XSSFSheet sheet, String text, int column, int row) {
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
