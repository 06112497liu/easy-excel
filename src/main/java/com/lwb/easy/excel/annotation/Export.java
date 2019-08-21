package com.lwb.easy.excel.annotation;

import com.lwb.easy.excel.enums.FileType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导出注解
 * </p>
 * 标记在方法上，用于确定哪个方法是导出方法，一般标记在controller方法上
 * @author liuweibo
 * @date 2019/8/14
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
public @interface Export {
    /**
     * 配置文件名称
     * </p>
     * 目前只支持yml类型文件的配置
     * @return
     */
    String value();
    /**
     * 配置文件的类型，默认是yml
     * @return
     */
    FileType type() default FileType.YML;
}
