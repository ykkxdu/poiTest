package com.poi.ykk.word;

import com.deepoove.poi.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.List;

/**
 * @Author:Yankaikai
 * @Description:
 * @Date:Created in 2018/10/4
 */
public class WordOperaterMethod {
    // 日志
    private static final Logger logger = LoggerFactory.getLogger(WordOperaterMethod.class);
    // 当前文档对象
    private XWPFDocument doc;
    public WordOperaterMethod(Path tplPath){
        try {
            FileInputStream in = new FileInputStream(tplPath.toFile());
            NiceXWPFDocument niceXWPFDocument = new NiceXWPFDocument(in);
            this.doc = niceXWPFDocument;
            operationWordMethod();
        }catch (Exception e)
        {
            e.printStackTrace();
        }
    }
    public void operationWordMethod(){
        // 得到文档中的所有元素，每一行就是一个元素。
        List<IBodyElement> bodyElements = doc.getBodyElements();
        int bodyElementSize = bodyElements.size();
        for(int i=0;i<bodyElementSize;i++)
        {
            // 得到当前段落的类型。
            BodyElementType type = bodyElements.get(i).getElementType();
            if (type != BodyElementType.PARAGRAPH) {
                continue;
            }
            //  获取当前段落的内容。
            XWPFParagraph paragraph = (XWPFParagraph) bodyElements.get(i);
            String text = paragraph.getText();
             if(text.contains("设计思路及要点"))
            {
                // 移除10号位置的段落。
                doc.removeBodyElement(i);
                bodyElementSize--;
                // 删除后指针已经移动，所有不需要++。
                i--;
            }else if(text.contains("主要施工工艺"))
            {
                //每个段落又由多个元素组成， 移除12位置paragraph段落的的一号元素。
                for(int m=0;m<paragraph.getRuns().size();m++)
                {
                    paragraph.removeRun(0);
                }
            }
        }
    }
}
