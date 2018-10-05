package com.poi.ykk.controller;

import com.poi.ykk.excel.ExcelOperaterMethod;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.nio.file.Path;
import java.nio.file.Paths;

import static javax.servlet.http.HttpServletResponse.SC_INTERNAL_SERVER_ERROR;
import static javax.servlet.http.HttpServletResponse.SC_NOT_FOUND;
import static javax.servlet.http.HttpServletResponse.SC_OK;

/**
 * @Author:Yankaikai
 * @Description:下载excel文件的接口
 * @Date:Created in 2018/10/4
 */
@Controller
@RequestMapping("api/download")
public class DownloadExcelController {
    @GetMapping("excel")
    public void exportExcel(HttpServletResponse response)
    {
        try {
            // 模板的名称
            String tplName="学生成绩表.xlsx";
            // 输出文件名称
            String outPutName="3年级二班"+tplName;
            // 模板路径
            Path templatePath=Paths.get("report","template",tplName);
            // 输出文件的路径
            Path finalOutputPath=Paths.get("report",outPutName);
            ExcelOperaterMethod utils=new ExcelOperaterMethod(templatePath,finalOutputPath);
            boolean success=utils.generateTable();
            if(success){
                response.setHeader("Content-Disposition",
                        "attachment; filename*=UTF-8''" + URLEncoder.encode(outPutName.toString(), "UTF-8")
                );
                utils.workbook.write(response.getOutputStream());
                utils.workbook.close(); // important
                response.setStatus(SC_OK);
            }else {
                response.setStatus(SC_INTERNAL_SERVER_ERROR);
            }

        }catch (Exception e)
        {
            response.setStatus(SC_NOT_FOUND);
        }

    }
}
