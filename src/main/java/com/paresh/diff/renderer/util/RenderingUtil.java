package com.paresh.diff.renderer.util;

import com.paresh.diff.cache.ClassMetadataCache;
import com.paresh.diff.dto.ClassMetadata;
import com.paresh.diff.dto.Diff;
import com.paresh.diff.dto.DiffResponse;
import com.paresh.diff.renderer.ExcelRendererConstants;
import com.paresh.diff.renderer.config.ExcelRenderingPreferences;
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
            headers = new ArrayList<>(classMetadata.getMethodDescriptions().size() + 1);
            headers.add(ExcelRendererConstants.DIFF_STATUS_HEADER);
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
                case DELETED:
                case NO_CHANGE:
                    consolidatedCollection.add(ClassMetadataCache.getInstance().getObjectFromIdentifier(diff.getIdentifier(), beforeCollection));
                    break;
            }
        }
        return consolidatedCollection;
    }

    public static String getTitle(ExcelRenderingPreferences renderingPreferences, Class collectionElementClass, ClassMetadata classMetadata) {
        String title = renderingPreferences.getTitle();
        if (title == null) {
            title = classMetadata.getClassDescription();
        }
        return title;
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
}
