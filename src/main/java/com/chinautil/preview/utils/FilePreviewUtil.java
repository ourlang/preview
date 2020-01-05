package com.chinautil.preview.utils;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.springframework.util.StringUtils;

import java.io.File;
import java.util.regex.Pattern;

/**
 * 这是一个工具类，主要是为了使Office2003-2007全部格式的文档(.doc|.docx|.xls|.xlsx|.ppt|.pptx)
 * 参考代码 ：https://github.com/ourlang/preview
 * @author 福小林
 */
public class FilePreviewUtil {

    private static final Log LOG = LogFactory.getLog(FilePreviewUtil.class);


    private FilePreviewUtil(){

    }

    /**
     * 使Office2003-2007全部格式的文档(.doc|.docx|.xls|.xlsx|.ppt|.pptx) 转化为pdf文件<br>
     * @param inputFilePath 源文件路径源文件路径
     * 源文件路径，如："e:/test.doc"
     * @return 转换后的pdf文件
     */
    public static File openOfficeToPdf(String inputFilePath) {
        return office2pdf(inputFilePath);
    }

    /**
     * 根据操作系统的名称，获取OpenOffice.org 4的安装目录<br>
     * 如我的OpenOffice.org 4安装在：C:\Program Files (x86)\OpenOffice 4
     * 请根据自己的路径做相应的修改
     * @return OpenOffice.org 4的安装目录
     */
    private static String getOfficeHome() {
        String osName = System.getProperty("os.name");
        System.out.println("操作系统名称:" + osName);
        if (Pattern.matches(Constant.LINUX_SYSTEM.getDescribe(), osName)) {
            return "/opt/openoffice.org4";
        } else if (Pattern.matches(Constant.WINDOWS_SYSTEM.getDescribe(), osName)) {
            return Constant.DEFAULT_INSTALL_PATH.getDescribe();
        } else if (Pattern.matches(Constant.MAC_SYSTEM.getDescribe(), osName)) {
            return "/Applications/OpenOffice.org.app/Contents/";
        }
        return Constant.DEFAULT_INSTALL_PATH.getDescribe();
    }

    /**
     * 连接OpenOffice.org 并且启动OpenOffice.org
     * @return OfficeManager对象
     */
    private static OfficeManager getOfficeManager() {
        DefaultOfficeManagerConfiguration config = new DefaultOfficeManagerConfiguration();
        // 设置OpenOffice.org 的安装目录
        String officeHome = getOfficeHome();
        config.setOfficeHome(officeHome);
        // 启动OpenOffice的服务
        OfficeManager officeManager = config.buildOfficeManager();
        officeManager.start();
        return officeManager;
    }

    /**
     * 把源文件转换成pdf文件
     * @param inputFile 文件对象
     * @param outputFilePath pdf文件路径
     * @param inputFilePath 源文件路径
     * @param converter 连接OpenOffice
     */
    private static File converterFile(File inputFile, String outputFilePath, String inputFilePath, OfficeDocumentConverter converter) {
        File outputFile = new File(outputFilePath);
        // 假如目标路径不存在,则新建该路径
        if (!outputFile.getParentFile().exists()) {
            outputFile.getParentFile().mkdirs();
        }
        converter.convert(inputFile, outputFile);
        System.out.println("文件：" + inputFilePath + "\n转换为\n目标文件：" + outputFile + "\n成功!");
        return outputFile;
    }

    /**
     * 使Office2003-2007全部格式的文档(.doc|.docx|.xls|.xlsx|.ppt|.pptx) 转化为pdf文件<br>
     * @param inputFilePath 源文件路径
     * 源文件路径，如："e:/test.doc"
     * 目标文件路径，如："e:/test_doc.pdf"
     * @return  对应的pdf文件
     */
    private static File office2pdf(String inputFilePath) {
        OfficeManager officeManager = null;
        try {
            if (StringUtils.isEmpty(inputFilePath)) {
                LOG.info("输入文件地址为空，转换终止!");
                return null;
            }
            File inputFile = new File(inputFilePath);
            // 转换后的文件路径 这里可以自定义转换后的路径,这里设置的和源文件同文件夹
            String outputFilePath = getOutputFilePath(inputFilePath);

            if (!inputFile.exists()) {
                LOG.info("输入文件不存在，转换终止!");
                return null;
            }
            // 获取OpenOffice的安装路劲
            officeManager = getOfficeManager();
            // 连接OpenOffice
            OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);

            return converterFile(inputFile, outputFilePath, inputFilePath, converter);
        } catch (Exception e) {
            LOG.error("转化出错!", e);
        } finally {
            // 停止openOffice
            if (officeManager != null) {
                officeManager.stop();
            }
        }
        return null;
    }

    /**
     * 获取输出文件路径 可自定义
     * @param inputFilePath 源文件的路径
     * @return 取输出pdf文件路径
     */
    private static String getOutputFilePath(String inputFilePath) {
        return inputFilePath.replaceAll("." + getPostfix(inputFilePath), ".pdf");
    }

    /**
     * 获取inputFilePath的后缀名，如："e:/test.doc"的后缀名为："doc"<br>
     * @param inputFilePath 源文件的路径
     * @return 源文件的后缀名
     */
    private static String getPostfix(String inputFilePath) {
        int index = inputFilePath.lastIndexOf('.');
        return inputFilePath.substring(index+1);
    }

}

