package com.poi.ykk.controller;

import com.deepoove.poi.XWPFTemplate;
import com.poi.ykk.excel.ExcelOperaterMethod;
import com.poi.ykk.word.WordOperaterMethod;
import com.poi.ykk.word.WordTableRenderPolicy;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ResourceLoader;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import static com.poi.ykk.utils.FileUtils.buildExportFileResponse;
import static com.poi.ykk.utils.TplUtils.replaceTemplate;

/**
 * @Author:Yankaikai
 * @Description:
 * @Date:Created in 2018/10/4
 */
@Controller
@RequestMapping("api/download")
public class DownLoadWordController {
    @Autowired
    private ResourceLoader resourceLoader;
    @GetMapping("word")
    public ResponseEntity exportWord()
    {
        try {
            // 模板名称
            String tplName="养护维修报告.docx";
            // 模板路径
            Path templatePath= Paths.get("report","template",tplName);
            // 输出报告名称
            String outputName="测试桥梁"+tplName;
            // 输出路径
            Path finalOutputPath=Paths.get("report",outputName);
            Map<String,Object> map=new HashMap<>();
            // 只是一个工具类，里面展示了如何对word进行操作：读取，删除，增加等。通常和poi-tl库结合起来使用。
            WordOperaterMethod util=new WordOperaterMethod(templatePath);
            // 将模版中的test进行替换
            map.put("test","word导出测试");
            // 根据中间模板初始化 poi-tl 模板,word通常结合poi-tl库进行操作使用。
            XWPFTemplate tpl = XWPFTemplate.compile(templatePath.toString());
            // 操作word中的表格，将数据填入。包括:合并单元格，添加表格。
            map.put("beginTable",new ExcelOperaterMethod().getStudentLists());
            tpl.registerPolicy("beginTable", new WordTableRenderPolicy());
            replaceTemplate(tpl, outputName, map);
            return buildExportFileResponse(resourceLoader,finalOutputPath,outputName);
        }catch (Exception e)
        {
            e.printStackTrace();
            return ResponseEntity.notFound().build();
        }
    }
}
