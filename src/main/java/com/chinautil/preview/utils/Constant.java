package com.chinautil.preview.utils;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 常量类
 * @author 福小林
 **/

@AllArgsConstructor
public enum  Constant {
    /**
     *使用的浏览器是linux系统
     */
    LINUX_SYSTEM(1,"Linux.*"),

    /**
     *使用的浏览器是windows系统
     */
    WINDOWS_SYSTEM(2,"Windows.*"),

    /**
     * 使用的浏览器是mac系统
     */
    MAC_SYSTEM(3,"Mac.*"),

    /**
     * 默认的OpenOffice安装路径 （默认为windows的）
     */
    DEFAULT_INSTALL_PATH(4,"C:\\Program Files (x86)\\OpenOffice 4");

    /**
     *编码（备用）
     */
    @Getter
    private int code;
    /**
     * 常量描述或常量赋值
     */
    @Getter
    private String describe;




}
