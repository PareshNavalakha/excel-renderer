package com.paresh.diff.renderer.util;

import com.paresh.diff.cache.ClassMetadataCache;
import com.paresh.diff.dto.ChangeType;
import com.paresh.diff.dto.ClassMetadata;
import com.paresh.diff.dto.Diff;
import com.paresh.diff.dto.DiffResponse;
import com.paresh.diff.renderer.ExcelRendererConstants;
import com.paresh.diff.renderer.config.ExcelRenderingPreferences;
import com.paresh.diff.renderer.config.ExcelStyles;
import com.paresh.diff.renderer.config.RenderingDefaults;
import com.paresh.diff.util.ReflectionUtil;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.*;

public class ExcelExport {

    private static final Logger logger = LoggerFactory.getLogger(ExcelExport.class);

    public static void exportToExcel(String fileName, Object before, Object after, DiffResponse diffResponse, ExcelRenderingPreferences renderingPreferences) {

        if (!CollectionUtils.isEmpty(diffResponse.getDiffs())) {
            diffResponse.getDiffs().removeIf(diff -> diff.getChangeType().equals(ChangeType.NO_CHANGE));

            try (Workbook workbook = new SXSSFWorkbook();) {

                generateReconSummarySheet(before, after, diffResponse, renderingPreferences, workbook);

                if (renderingPreferences.isSourceDumpRequired()) {
                    generateSourceDumps(before, after, diffResponse, renderingPreferences, workbook);
                }

                FileOutputStream out = new FileOutputStream(fileName);
                workbook.write(out);
                out.close();
                logger.info("Done saving file {}", fileName);
            } catch (IOException e) {
                logger.error("Exception while exporting excel", e);
            }

        } else {
            //TODO to check if Excel sheet really required.
            logger.info("No diffs found.");
        }
    }

    private static void generateSourceDumps(Object before, Object after, DiffResponse diffResponse, ExcelRenderingPreferences renderingPreferences, Workbook workbook) {
        Class collectionElementClass = ReflectionUtil.getCollectionElementClass(before, after);
        ClassMetadata classMetadata = diffResponse.getClassMetaDataMap().get(collectionElementClass);
        List<String> headers = RenderingUtil.getHeaders(classMetadata);

        if (before != null) {
            generateDump((Collection) before, renderingPreferences.getExcelStyles(), workbook, classMetadata, headers, renderingPreferences.getBeforeSheetName());
        }

        if (after != null) {
            generateDump((Collection) after, renderingPreferences.getExcelStyles(), workbook, classMetadata, headers, renderingPreferences.getAfterSheetName());
        }
    }

    private static void generateDump(Collection collection, ExcelStyles excelStyles, Workbook workbook, ClassMetadata classMetadata, List<String> headers, String sheetName) {
        Sheet sheet = createSheet(workbook, sheetName);
        int rowNumber = 1;
        generateHeaderRows(workbook, sheet, headers, rowNumber++, excelStyles);
        if (!CollectionUtils.isEmpty(collection)) {
            for (Object object : collection) {
                List<Method> methods = classMetadata.getMethods();
                ChangeType attributeChangeType;
                Diff relevantDiff;
                for (int columnIndex = 0; columnIndex < methods.size(); columnIndex++) {
                    attributeChangeType = ChangeType.NO_CHANGE;
                    generateCellContent(sheet, rowNumber, columnIndex, RenderingUtil.getAttributeObject(methods.get(columnIndex), object), "", excelStyles.withWorkBook(workbook).getUnchangedStyle());
                }
                rowNumber++;
            }
        }
    }

    private static void generateReconSummarySheet(Object before, Object after, DiffResponse diffResponse, ExcelRenderingPreferences renderingPreferences, Workbook workbook) {

        Collection beforeCollection = (Collection) before;
        Collection afterCollection = (Collection) after;

        Map<Object, Object> beforeIdentifierMap = new HashMap<>(beforeCollection.size());
        Map<Object, Object> afterIdentifierMap = new HashMap<>(afterCollection.size());

        beforeCollection.stream().forEach(element ->
                beforeIdentifierMap.put(ClassMetadataCache.getInstance().getIdentifier(element), element));

        afterCollection.stream().forEach(element ->
                afterIdentifierMap.put(ClassMetadataCache.getInstance().getIdentifier(element), element));

        List consolidatedList = RenderingUtil.getConsolidatedCollection(before, beforeIdentifierMap, after, afterIdentifierMap, diffResponse);

        if (!CollectionUtils.isEmpty(consolidatedList)) {
            int rowNumber = 0;
            Class collectionElementClass = ReflectionUtil.getCollectionElementClass(before, after);
            ClassMetadata classMetadata = diffResponse.getClassMetaDataMap().get(collectionElementClass);

            List<String> headers = new ArrayList<>();

            headers.addAll(RenderingUtil.getMetaDataHeaders(renderingPreferences));
            headers.addAll(RenderingUtil.getHeaders(classMetadata));

            Sheet sheet = createSheet(workbook, renderingPreferences.getSheetName());

            ExcelStyles excelStyles = renderingPreferences.getExcelStyles();

            generateTitleRow(workbook, sheet, RenderingUtil.getTitle(renderingPreferences, collectionElementClass, classMetadata), excelStyles, headers.size() - 1);
            generateHeaderRows(workbook, sheet, headers, ++rowNumber, excelStyles);

            ChangeType changeType;
            String changeTypeText = null;

            int outerIndex = 0;
            for (Diff diff : diffResponse.getDiffs()) {
                int columnIndex = 0;
                changeType = diff.getChangeType();
                switch (changeType) {
                    case ADDED:
                        changeTypeText = RenderingDefaults.getNewDescription(renderingPreferences.getRenderingMode(), renderingPreferences.getSource2());
                        break;
                    case UPDATED:
                        changeTypeText = RenderingDefaults.getUpdatedDescription(renderingPreferences.getRenderingMode());
                        break;
                    case DELETED:
                        changeTypeText = RenderingDefaults.getDeletedDescription(renderingPreferences.getRenderingMode(), renderingPreferences.getSource2());
                        break;
                    case NO_CHANGE:
                        changeTypeText = ExcelRendererConstants.BLANK;
                        break;
                    default:
                        changeTypeText = ExcelRendererConstants.BLANK;
                        break;

                }
                if (renderingPreferences.isChangeTypeHeaderRequired()) {
                    generateCellContent(sheet, ++rowNumber, columnIndex, changeTypeText, ExcelRendererConstants.BLANK, getCellStyle(workbook, changeType, renderingPreferences.getExcelStyles()));
                    columnIndex++;
                }

                if (renderingPreferences.isSummaryOfChangeHeaderRequired()) {
                    generateCellContent(sheet, rowNumber, columnIndex, RenderingUtil.getSummaryOfChange(diff.getChildDiffs()), ExcelRendererConstants.BLANK, renderingPreferences.getExcelStyles().withWorkBook(workbook).getUnchangedStyle());
                    columnIndex++;
                }

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
                    generateCellContent(sheet, rowNumber, innerIndex + columnIndex, RenderingUtil.getAttributeObject(methods.get(innerIndex), consolidatedList.get(outerIndex)),
                            getComment(diff,
                                    methods.get(innerIndex),
                                    attributeChangeType,
                                    beforeIdentifierMap,
                                    afterIdentifierMap,
                                    renderingPreferences), getCellStyle(workbook, attributeChangeType, excelStyles));
                }
                outerIndex++;
            }
        }
    }

    private static String getComment(Diff classDiff, Method method, ChangeType changeType, Map<Object, Object> beforeIdentifierMap, Map<Object, Object> afterIdentifierMap, ExcelRenderingPreferences renderingPreferences) {
        String response = null;
        Object afterObject = null;
        Object beforeObject = null;

        switch (changeType) {
            case UPDATED:
                beforeObject = ClassMetadataCache.getInstance().getCorrespondingObjectMatchingIdentifier(classDiff.getIdentifier(), beforeIdentifierMap);
                if (renderingPreferences.getSource1() != null) {
                    response = renderingPreferences.getSource1() + " : " + ReflectionUtil.getMethodResponse(method, beforeObject);
                } else {
                    response = "Original : " + ReflectionUtil.getMethodResponse(method, beforeObject);

                }
                break;
            case DELETED:
                afterObject = ClassMetadataCache.getInstance().getCorrespondingObjectMatchingIdentifier(classDiff.getIdentifier(), afterIdentifierMap);
                if (renderingPreferences.getSource2() != null) {
                    response = renderingPreferences.getSource2() + " : " + ReflectionUtil.getMethodResponse(method, afterObject);
                } else {
                    response = "Original : " + ReflectionUtil.getMethodResponse(method, afterObject);
                }
                break;
            case NO_CHANGE:
            case ADDED:
            default:
                break;
        }
        return response;
    }


    private static Sheet createSheet(Workbook workbook, String sheetName) {
        return workbook.createSheet(sheetName);
    }

    private static void generateTitleRow(Workbook workbook, Sheet sheet, String title, ExcelStyles excelStyles, int headerCount) {
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(excelStyles.withWorkBook(workbook).getTitleStyle());
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headerCount));
    }

    private static void generateHeaderRows(Workbook workbook, Sheet sheet, List<String> headers, int row, ExcelStyles excelStyles) {
        Row headerRow = sheet.createRow(row);
        headerRow.setHeightInPoints(40);
        Cell headerCell;
        for (int index = 0; index < headers.size(); index++) {
            headerCell = headerRow.createCell(index);
            headerCell.setCellValue(headers.get(index));
            headerCell.setCellStyle(excelStyles.withWorkBook(workbook).getHeaderStyle());
        }
    }

    private static CellStyle getCellStyle(Workbook workbook, ChangeType changeType, ExcelStyles excelStyles) {
        switch (changeType) {
            case ADDED:
                return excelStyles.withWorkBook(workbook).getNewStyle();
            case UPDATED:
                return excelStyles.withWorkBook(workbook).getModifiedStyle();
            case DELETED:
                return excelStyles.withWorkBook(workbook).getDeletedStyle();
            case NO_CHANGE:
                return excelStyles.withWorkBook(workbook).getUnchangedStyle();
            default:
                return excelStyles.withWorkBook(workbook).getUnchangedStyle();
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
            cell.setCellComment(generateCellComment((SXSSFSheet) sheet, comment, column, row));
        }
        cell.setCellStyle(cellStyle);
    }

    private static Comment generateCellComment(SXSSFSheet sheet, String text, int column, int row) {
        SXSSFDrawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, column, row, column + 3, row + 3);
        Comment comment = drawing.createCellComment(anchor);

        // set text in the comment
        comment.setString(new XSSFRichTextString(text));

        //set comment author.
        //you can see it in the status bar when moving mouse over the commented cell
        comment.setAuthor("object-diff");

        return comment;
    }
}
