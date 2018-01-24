package com.paresh.diff.renderer.util;

import com.paresh.diff.cache.ClassMetadataCache;
import com.paresh.diff.dto.ChangeType;
import com.paresh.diff.dto.ClassMetadata;
import com.paresh.diff.dto.Diff;
import com.paresh.diff.dto.DiffResponse;
import com.paresh.diff.renderer.config.ExcelRenderingPreferences;
import com.paresh.diff.renderer.config.RenderingDefaults;
import com.paresh.diff.util.ReflectionUtil;
import org.apache.commons.collections4.CollectionUtils;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class RenderingUtil {

    //TODO Implement ordering
    public static List<String> getHeaders(ClassMetadata classMetadata) {
        List<String> headers = null;
        if (!CollectionUtils.isEmpty(classMetadata.getMethodDescriptions())) {
            headers = new ArrayList<>(classMetadata.getMethodDescriptions().size());
            headers.addAll(classMetadata.getMethodDescriptions());
        }
        return headers;
    }

    public static Object getAttributeObject(Method method, Object o) {
        return ReflectionUtil.getMethodResponse(method, o);
    }

    public static List getConsolidatedCollection(Object before, Object after, DiffResponse diffResponse) {
        List consolidatedCollection = new ArrayList(diffResponse.getDiffs().size());
        Collection beforeCollection = (Collection) before;
        Collection afterCollection = (Collection) after;

        for (Diff diff : diffResponse.getDiffs()) {
            switch (diff.getChangeType()) {
                case ADDED:
                case UPDATED:
                    consolidatedCollection.add(ClassMetadataCache.getInstance().getObjectFromIdentifier(diff.getIdentifier(), afterCollection));
                    break;
                case NO_CHANGE:
                case DELETED:
                    consolidatedCollection.add(ClassMetadataCache.getInstance().getObjectFromIdentifier(diff.getIdentifier(), beforeCollection));
                    break;
                default:
                    break;
            }
        }
        return consolidatedCollection;
    }

    public static String getTitle(ExcelRenderingPreferences renderingPreferences, Class collectionElementClass, ClassMetadata classMetadata) {
        String title = renderingPreferences.getTitle();
        if (title == null) {
            title = classMetadata.getClassDescription() + RenderingDefaults.getTitlePostfix(renderingPreferences.getRenderingMode());
        }
        return title;
    }

    private static String getChangeSummaryForChangeType(Collection<Diff> childDiffs, ChangeType changeType) {
        boolean found = false;
        StringBuilder response = new StringBuilder();
        if (!CollectionUtils.isEmpty(childDiffs)) {
            for (Diff diff : childDiffs) {
                if (changeType.equals(diff.getChangeType())) {
                    response.append(found ? ", " + diff.getFieldDescription() : diff.getFieldDescription());
                    found = true;
                }
            }
        }
        if (found) {
            return changeType + " fields: " + response.toString() + ".";
        } else {
            return response.toString();
        }
    }

    public static String getSummaryOfChange(Collection<Diff> childDiffs) {
        StringBuilder response = new StringBuilder();
        if (childDiffs != null) {
            childDiffs.removeIf(diff -> diff.getChangeType().equals(ChangeType.NO_CHANGE));
            if (!CollectionUtils.isEmpty(childDiffs)) {
                response.append(getChangeSummaryForChangeType(childDiffs, ChangeType.UPDATED));
                response.append(getChangeSummaryForChangeType(childDiffs, ChangeType.DELETED));
                response.append(getChangeSummaryForChangeType(childDiffs, ChangeType.ADDED));
            }
        }
        return response.toString();
    }


    public static Diff getRelevantDiff(Collection<Diff> childDiffs, Method method, ClassMetadata classMetadata) {
        if (!CollectionUtils.isEmpty(childDiffs)) {
            String description = classMetadata.getMethodDescriptions().get(classMetadata.getMethods().indexOf(method));
            for (Diff diff : childDiffs) {
                if (description.equals(diff.getFieldDescription())) {
                    return diff;
                }
            }
        }
        return null;
    }

    public static List<String> getMetaDataHeaders(ExcelRenderingPreferences renderingPreferences) {
        List<String> metaDataHeaders = new ArrayList<>();
        if (renderingPreferences.isChangeTypeHeaderRequired()) {
            metaDataHeaders.add("Change Type");
        }
        if (renderingPreferences.isSummaryOfChangeHeaderRequired()) {
            metaDataHeaders.add("Summary of changes");
        }
        return metaDataHeaders;
    }
}
