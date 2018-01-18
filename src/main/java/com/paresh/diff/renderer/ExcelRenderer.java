package com.paresh.diff.renderer;


import com.paresh.diff.dto.DiffResponse;
import com.paresh.diff.renderer.config.ExcelRenderingPreferences;
import com.paresh.diff.renderer.util.ExcelExport;
import com.paresh.diff.renderer.util.FileUtilty;
import com.paresh.diff.util.ReflectionUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Collection;
import java.util.Iterator;

public class ExcelRenderer implements Renderer {

    private final Logger logger = LoggerFactory.getLogger(getClass());

    @Override
    public void render(Object before, Object after, DiffResponse diffResponse, RenderingPreferences renderingPreferences) {
        checkIfObjectsAreInstancesOfCollections(before, after);
        String fileName = (String) renderingPreferences.getPreference(ExcelRendererConstants.EXPORT_FILE_NAME);

        if (fileName == null) {
            fileName = FileUtilty.getTemporaryLocation() + FileUtilty.getSeperator() + generateFileName(before, after, diffResponse);
            System.setProperty(ExcelRendererConstants.EXPORT_FILE_NAME, fileName);
        }

        ExcelExport.exportToExcel(fileName, before, after, diffResponse, (ExcelRenderingPreferences) renderingPreferences);

    }

    private String generateFileName(Object before, Object after, DiffResponse diffResponse) {
        Class clazz = ReflectionUtil.getCollectionElementClass(before, after);
        return diffResponse.getClassMetaDataMap().get(clazz).getClassDescription() + ExcelRendererConstants.EXTENSION;
    }

    @Override
    public void render(Object before, Object after, DiffResponse diffResponse) {
        render(before, after, diffResponse, new ExcelRenderingPreferences());
    }

    private void checkIfObjectsAreInstancesOfCollections(Object before, Object after) {
        String errorMessage = "Object is not an instance of collection.";
        if (before != null) {
            if (!isAComplexCollection(before)) {
                logger.error(errorMessage);
                throw new IllegalArgumentException(errorMessage);
            }
        } else if (after != null) {
            if (!isAComplexCollection(after)) {
                logger.error(errorMessage);
                throw new IllegalArgumentException(errorMessage);
            }
        }
    }

    private boolean isAComplexCollection(Object object) {
        if (ReflectionUtil.isInstanceOfCollection(object)) {
            Collection collection = (Collection) object;
            if (collection != null && !collection.isEmpty()) {
                Iterator iterator = collection.iterator();
                return !ReflectionUtil.isBaseClass(iterator.next().getClass());
            }
        }
        return false;
    }
}
