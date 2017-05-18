package com.piedra.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导出字段注解
 * 对于 浮点数，日期 等其他特殊字段，可以在要接收导入数据的实体中加上一个set方法来接收字符串，在将字符串转换成其他数据
 * @author linwb
 * @since  2017-05-18
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelImport {

    /**
     * 列的顺序 以 0,1,2,3,4 ... 表示
     */
    String colIndex() ;
}
