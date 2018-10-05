package com.poi.ykk.word;

import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.poi.ykk.entity.Student;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.List;

import static com.poi.ykk.utils.TplUtils.*;

/**
 * @Author:Yankaikai
 * @Description:
 * @Date:Created in 2018/10/4
 */
public class WordTableRenderPolicy extends DynamicTableRenderPolicy{
    @Override
    public void render(XWPFTable xwpfTable, Object o) {
        List<Student> data= (List<Student>) o;
        List<XWPFTableRow> rows = xwpfTable.getRows();
        int rowSize = rows.size();
        int rowNo = 1;
        // 数据列向量,用于动态的添加数据
        int initialColSize = 8;
        // 序号
        int count=1;
        // 创建行数和列数
        for(int i=2;i<data.size()+2;i++)
        {
            XWPFTableRow row=xwpfTable.insertNewTableRow(i);
            addCells(row,initialColSize);
            rowSize++;
        }
        for(int i=0;i<data.size();i++)
        {
            // 表格不够，需要重新添加
            XWPFTableRow row = rows.get(rowNo);
            // 得到每一行中的元素。
            List<XWPFTableCell> cells = row.getTableCells();
            if (cells.get(0).getText().equals("{{beginTable}}")) {
                for (XWPFParagraph paragraph : cells.get(0).getParagraphs()) {
                    if(paragraph.getRuns().get(0).getText(0).equals("{{beginTable}}"))
                    {
                        paragraph.removeRun(0);
                        break;
                    }
                }
            }

            cells.get(0).setText(String.valueOf(count));
            cells.get(1).setText(data.get(i).getName());
            cells.get(2).setText(data.get(i).getNum());
            cells.get(3).setText(data.get(i).getClassLevel());
            cells.get(4).setText(data.get(i).getMath());
            cells.get(5).setText(data.get(i).getChinese());
            cells.get(6).setText(data.get(i).getEnglish());
            int sum=Integer.parseInt(data.get(i).getMath())+
                    Integer.parseInt(data.get(i).getChinese())+
                    Integer.parseInt(data.get(i).getEnglish());
            cells.get(7).setText(String.valueOf(sum));
            count++;
            rowNo++;
        }
        // 在最后一行添加班主任
        XWPFTableRow row = rows.get(rowNo);
        List<XWPFTableCell> cells = row.getTableCells();
        cells.get(0).setText("班主任:");
        // 进行合并单元格,共有三个参数，操作的表格，合并的列数，起始行与终止行。
        mergeCellsVertically(xwpfTable,3,1,3);
        mergeCellsVertically(xwpfTable,3,4,5);
        // 水平合并最后一行
        mergeCellsHorizontally(xwpfTable,rowSize-1,0,7);

    }
}
