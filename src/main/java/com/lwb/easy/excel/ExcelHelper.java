package com.lwb.easy.excel;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory;
import com.lwb.easy.excel.annotation.Export;
import com.lwb.easy.excel.exception.ExcelException;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Stream;

import static com.lwb.easy.excel.constant.Constant.*;

/**
 * excel文件生成工具类
 * @author liuweibo
 * @date 2019/8/20
 */
public class ExcelHelper {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelHelper.class);

    /**
     * 从方法调用栈中获取标注有制定注解的方法
     * @param type       注解类型
     * @param stackTrace 方法调用栈
     * @return 方法栈
     */
    static <T extends Annotation> Method getMethod(Class<T> type, StackTraceElement... stackTrace) {
        return
            Arrays.stream(stackTrace)
                .map(trace -> {
                    try {
                        Class<?> clazz = Class.forName(trace.getClassName());
                        return
                            Optional.ofNullable(clazz.getDeclaredMethods())
                                .filter(ArrayUtils::isNotEmpty)
                                .map(methods -> Stream.of(methods)
                                    .filter(method -> method.getAnnotation(type) != null)
                                    .findAny()
                                    .orElse(null)
                                )
                                .orElse(null);
                    } catch (ClassNotFoundException e) {
                        LOGGER.error(e.getMessage(), e);
                    }
                    return null;
                })
                .filter(Objects::nonNull)
                .findAny()
                .orElseThrow(() -> new ExcelException("没有找到Export标记的方法!"));
    }

    /**
     * 获取字段值
     * @param obj       对象
     * @param fieldName 字段名称
     * @return 字段值，转换成了String
     */
    public static String getFieldValue(Object obj, String fieldName) throws NoSuchFieldException, IllegalAccessException {
        if (obj == null || StringUtils.isEmpty(fieldName)) {
            return EMPTY;
        }
        // 如果传入对象是map 直接获取key值
        if (obj instanceof Map) {
            return ((Map) obj).containsKey(fieldName) ? (String) ((Map) obj).get(fieldName) : EMPTY;
        }
        // 支持获取嵌套对象的值（例如：user.role.name，表示获取user对象中嵌套对象role的name字段的值）
        if (fieldName.contains(POINT)) {
            int i = fieldName.indexOf(POINT);
            String currentFieldName = fieldName.substring(0, i);
            String nextFieldName = fieldName.substring(i + 1, fieldName.length());
            Field field = obj.getClass().getDeclaredField(currentFieldName);
            if (!field.isAccessible()) {
                field.setAccessible(true);
            }
            Object o = field.get(obj);
            if (field.isAccessible()) {
                field.setAccessible(false);
            }
            // 当前字段为null，不在向下获取值
            if (o == null) {
                return EMPTY;
            }
            return getFieldValue(o, nextFieldName);
        } else {
            return formatFieldValue(obj, fieldName);
        }
    }

    /**
     * 格式化字段的值
     * </p>
     * 日期字段根据JsonFormat注解的样式格式化，没有设置则使用相关默认的格式
     * @param obj       对象
     * @param fieldName 字段名称
     * @return 格式化后的值
     */
    private static String formatFieldValue(Object obj, String fieldName) throws NoSuchFieldException, IllegalAccessException {
        Field field = obj.getClass().getDeclaredField(fieldName);
        if (!field.isAccessible()) {
            field.setAccessible(true);
        }
        Object o = field.get(obj);
        if (field.isAccessible()) {
            field.setAccessible(false);
        }

        return Optional.of(o)
            .filter(ExcelHelper::isDate)
            .map(d -> {
                String pattern = Optional.ofNullable(field.getAnnotation(JsonFormat.class))
                    .map(JsonFormat::pattern)
                    .orElseGet(() -> {
                        try {
                            PropertyDescriptor descriptor = new PropertyDescriptor(fieldName, o.getClass());
                            return Optional.ofNullable(descriptor.getWriteMethod())
                                .map(m -> m.getAnnotation(JsonFormat.class))
                                .map(JsonFormat::pattern)
                                .orElse(null);
                        } catch (IntrospectionException e) {
                            LOGGER.error(e.getMessage(), e);
                            return null;
                        }
                    });
                return dateFormat(o, pattern);
            })
            .orElse(String.valueOf(o));
    }

    /**
     * 对象是不是日期对象
     * @param obj
     * @return
     */
    private static boolean isDate(Object obj) {
        return (obj instanceof Date) ||
            (obj instanceof LocalDateTime) ||
            (obj instanceof LocalDate) ||
            (obj instanceof LocalTime);
    }

    /**
     * 转换日期格式
     * @param date    具体对象
     * @param pattern 格式化格式
     * @return 格式化后的值
     */
    private static String dateFormat(Object date, String pattern) {
        if (date instanceof Date) {
            return DateFormatUtils.format((Date) date, pattern);
        } else if (date instanceof LocalTime) {
            return
                DateTimeFormatter.ofPattern(pattern == null ? HH_MM_SS : pattern).format((LocalTime) date);
        } else if (date instanceof LocalDate) {
            return
                DateTimeFormatter.ofPattern(pattern == null ? YYYY_MM_DD : pattern).format((LocalDate) date);
        } else if (date instanceof LocalDateTime) {
            return
                DateTimeFormatter.ofPattern(pattern == null ? YYYY_MM_DD_HH_MM_SS : pattern).format((LocalDateTime) date);
        }
        return String.valueOf(date);
    }

    /**
     * 解析yml文件
     * </p>
     * 解析成ExcelConfig，用于后续初始化excel
     * @param method 被某个注解标记的方法
     * @return
     */
    public static ExcelConfig parseYml(Method method) {
        try {
            Class<?> clazz = method.getDeclaringClass();
            Export exportConfig = method.getAnnotation(Export.class);
            ObjectMapper mapper = new ObjectMapper(new YAMLFactory());
            return mapper.readValue(
                clazz.getResourceAsStream(exportConfig.value()),
                ExcelConfig.class
            );
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelException(e.getMessage());
        }
    }

    /**
     * 获取当前excel导出的配置文件
     * @return
     */
    public static ExcelConfig parseConfig() {
        return parseYml(getMethod(Export.class, Thread.currentThread().getStackTrace()));
    }

}
