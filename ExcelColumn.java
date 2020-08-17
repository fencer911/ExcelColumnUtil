package cn.fencer911;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


/**
 * 数值型的栏位只能使用Double
 * @author https://github.com/fencer911/ExcelColumnUtil
 * @version 1.0, Created at 2020年2月20日
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {
    /**
     * 顺序 default 100
     * 
     * @return index
     */
    String value();

    /**
     * 当值为null时要显示的值 default StringUtils.EMPTY
     * 
     * @return defaultValue
     */
    String defaultValue() default "";

    /**
     * 用于验证
     * 
     * @return valid
     */
    Valid valid() default @Valid();

    @Retention(RetentionPolicy.RUNTIME)
    @Target(ElementType.FIELD)
    @interface Valid {
        /**
         * 必须与in中String相符,目前仅支持String类型
         * 
         * @return e.g. {"key","value"}
         */
        String[] in() default {};

        /**
         * 是否允许为空,用于验证数据 default true
         * 
         * @return allowNull
         */
        boolean allowNull() default false;

        /**
         * Apply a "greater than" constraint to the named property
         * 
         * @return gt
         */
        double gt() default Double.NaN;

        /**
         * Apply a "less than" constraint to the named property
         * @return lt
         */
        double lt() default Double.NaN;

        /**
         * Apply a "greater than or equal" constraint to the named property
         * 
         * @return ge
         */
        double ge() default Double.NaN;

        /**
         * Apply a "less than or equal" constraint to the named property
         * 
         * @return le
         */
        double le() default Double.NaN;
    }
}
