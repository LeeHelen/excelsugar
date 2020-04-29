package com.kyle.excel.util;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.beanutils.PropertyUtils;

/**
 * @package: com.kyle.excel.util
 * @className: BeanUtil
 * @author: Kyle.Y.Li
 * @since 1.0.0 2020-04-4/29/2020 16:47
 */
public class BeanUtil {
    /**
     * bean转换为Map<String, String>
     *
     * @param obj
     * @return
     */
    public static Map<String, String> objectToMapStr(Object obj) {
        Map<String, String> map = new LinkedHashMap<>();
        if (obj == null) {
            return map;
        }
        try {
            map = BeanUtils.describe(obj);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
        return map;
    }

    /**
     * bean转换为Map<?,?>
     *
     * @param obj
     * @return Map<?,?>
     */
    public static Map<?, ?> objectToMapAny(Object obj) {
        Map<?, ?> map = new LinkedHashMap<>();
        if (obj == null) {
            return map;
        }
        return new org.apache.commons.beanutils.BeanMap(obj);
    }

    /**
     * bean转换为Map<String, Object>
     *
     * @param obj
     * @return
     */
    public static Map<String, Object> objectToMap(Object obj) {
        Map<String, Object> map = new LinkedHashMap<>();
        if (obj == null) {
            return map;
        }
        try {
            map = PropertyUtils.describe(obj);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
        return map;
    }

	/*public static Map<String, Object> getBeanMap(Object bean) {
	    Map<String, Object> beanMap = new HashMap<String, Object>();
	    BeanWrapper beanWrapper = new BeanWrapper(BeanWrapperContext.create(bean.getClass()));
	    for(String propertyName : beanWrapper.getPropertyNames())
	        beanMap.put(propertyName, beanWrapper.getValue(propertyName));
	    return beanMap;
	}*/

    /**
     * map转换为bean
     *
     * @param map
     * @param beanClass
     * @return bean
     * @throws Exception
     */
    public static <T> T mapToObject(Map<String, Object> map, Class<T> beanClass) {
        T obj = null;
        if (map == null) {
            return obj;
        }
        try {
            BeanUtils.populate(beanClass, map);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }
        return obj;
    }

    /**
     * 获取字段的注解
     *
     * @param field 字段
     */
    public static <T extends Annotation> T getAnnotation(Field field, Class<T> annotationClass) {
        if(field == null)
            return null;

        return field.getAnnotation(annotationClass);
    }
}
