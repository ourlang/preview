package com.chinautil.preview.controller;

import com.chinautil.preview.utils.FilePreviewUtil;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.net.URLEncoder;

/**
 * 测试文件在线预览的控制层
 * service和dao层这里省略
 * @author 福小林
 **/
@RestController
public class FilePreviewController {
    private static final Log LOG = LogFactory.getLog(FilePreviewController.class);


    /**
     * 在线预览文件的方法
     * 测试地址 http://localhost:8099/chinaUtil/getFile?filePath=E:/file/test.doc
     * @param response 响应的response对象
     * @param request 请求消息Request对象
     */

    @ResponseBody
    @RequestMapping("/getFile")
    public void readFile(HttpServletResponse response , HttpServletRequest request){
        try {
            //源文件的路径
            String filePath = request.getParameter("filePath");
            File file = FilePreviewUtil.openOfficeToPdf(filePath);
            //把转换后的pdf文件响应到浏览器上面
            FileInputStream fileInputStream = new FileInputStream(file);
            BufferedInputStream br = new BufferedInputStream(fileInputStream);
            byte[] buf = new byte[1024];
            int length;
            // 清除首部的空白行。非常重要
            response.reset();
            //设置调用浏览器嵌入pdf模块来处理相应的数据。
            response.setContentType("application/pdf");
            response.setHeader("Content-Disposition","inline; filename=" + URLEncoder.encode(file.getName(), "UTF-8"));
            OutputStream out = response.getOutputStream();
            //写文件
            while ((length = br.read(buf)) != -1){
                out.write(buf, 0, length);
            }
            br.close();
            out.close();
        }catch (Exception e){
            LOG.error("在线预览异常",e);
        }

    }
}
