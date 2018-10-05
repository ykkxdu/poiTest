package com.poi.ykk.excel;

import com.poi.ykk.entity.Student;
import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static com.poi.ykk.utils.FileUtils.isExist;

/**
 * @Author:Yankaikai
 * @Description:excel表格操作方法:填写数据，合并单元格，字体样式以及边框
 * @Date:Created in 2018/10/4
 */
@Data
public class ExcelOperaterMethod {
    // 模板路径
    public Path templatePath;
    // 输出路径
    public  Path finalOutputPath;
    // xlsx 操作类
    public XSSFWorkbook workbook;
    public List<Student> studentLists;
    // 日志
    private static final Logger logger = LoggerFactory.getLogger(ExcelOperaterMethod.class);
    // 默认单元格样式
    private static CellStyle defaultCellStyle = null;
    public ExcelOperaterMethod()
    {
        studentLists=new ArrayList<>();
        initialStudentData();
    }
    public ExcelOperaterMethod(Path templatePath,Path finalOutputPath)
    {
        this.templatePath=templatePath;
        this.finalOutputPath=finalOutputPath;
        studentLists=new ArrayList<>();
    }
    /**
    * @Author：Yankaikai
    * @Description:产生表格
    * @Date: 2018/10/4
    * @param: 无
    */
    public boolean generateTable(){
        try {
            if(isExist(finalOutputPath)) {
                Files.delete(finalOutputPath);
            }
            Files.copy(templatePath,finalOutputPath);
            workbook=new XSSFWorkbook(finalOutputPath.toFile());
            // 默认边框及字体
            defaultCellStyle=generateDefaultCellStyle(workbook);
            // 初始化构件目录树表头
            XSSFSheet sheet=workbook.getSheetAt(0);
            initialStudentData();
            initialTable(sheet);
            writeTableData(sheet);
            return true;
        }catch (Exception e)
        {
            logger.warn(e.getMessage());
            e.printStackTrace();
            return false;
        }
    }

    public void writeTableData(XSSFSheet sheet)
    {
        // 序号
        int count=1;
        int initialRow=1;
        for(int i=0;i<studentLists.size();i++)
        {
            int initialCell=0;
            // 得到创建的行进行填写数据
            XSSFRow row=sheet.getRow(initialRow);
            row.getCell(initialCell).setCellValue(count);
            initialCell++;
            row.getCell(initialCell).setCellValue(studentLists.get(i).getName());
            initialCell++;
            row.getCell(initialCell).setCellValue(studentLists.get(i).getNum());
            initialCell++;
            row.getCell(initialCell).setCellValue(studentLists.get(i).getClassLevel());
            initialCell++;
            row.getCell(initialCell).setCellValue(studentLists.get(i).getMath());
            initialCell++;
            row.getCell(initialCell).setCellValue(studentLists.get(i).getChinese());
            initialCell++;
            row.getCell(initialCell).setCellValue(studentLists.get(i).getEnglish());
            initialCell++;
            int sum=Integer.parseInt(studentLists.get(i).getMath())+
                    Integer.parseInt(studentLists.get(i).getChinese())+
                    Integer.parseInt(studentLists.get(i).getEnglish());
            row.getCell(initialCell).setCellValue(sum);
            initialCell++;
            initialRow++;
            count++;
        }
        XSSFRow row=sheet.getRow(initialRow);
        XSSFCell cell=row.getCell(0);
        cell.setCellValue("班主任:");
        cell.setCellStyle(defaultCellStyle);
        // 合并单元格
        mergeUnit(sheet);
    }
    // 现在需要将第3列1-3行进行合并。第3列的4,5行合并。
    // 只要记住合并的行和列，都可以进行操作。
    public void mergeUnit(XSSFSheet sheet)
    {
        // 四个参数:n-m行，x-y列
        //sheet.addMergedRegion(new CellRangeAddress(2,2,3,5));
        sheet.addMergedRegion(new CellRangeAddress(1,3,3,3));
        sheet.addMergedRegion(new CellRangeAddress(4,5,3,3));
        sheet.addMergedRegion(new CellRangeAddress(studentLists.size()+1,studentLists.size()+1,0,7));
    }
    // 初始化表格
    public void initialTable(XSSFSheet sheet)
    {
        // 从第二行开始
        int initialRow=1;
        for(int i=0;i<studentLists.size()+1;i++)
        {
            // 第一步:必须创建行
            XSSFRow row=sheet.createRow(initialRow);
            // 第二步:根据创建的行创造8列
            for(int m=0;m<8;m++)
            {
                XSSFCell cell=row.createCell(m);
                cell.setCellStyle(defaultCellStyle);
            }
            initialRow++;
        }

    }
    // 初始化实体
    public void initialStudentData()
    {
        Student stu1=new Student();
        stu1.setNum("123");
        stu1.setName("严凯凯");
        stu1.setClassLevel("二年级");
        stu1.setMath("80");
        stu1.setChinese("90");
        stu1.setEnglish("70");
        Student stu2=new Student();
        stu2.setNum("124");
        stu2.setName("王浩");
        stu2.setClassLevel("二年级");
        stu2.setMath("70");
        stu2.setChinese("80");
        stu2.setEnglish("90");
        Student stu3=new Student();
        stu3.setNum("125");
        stu3.setName("pant");
        stu3.setClassLevel("三年级");
        stu3.setMath("60");
        stu3.setChinese("67");
        stu3.setEnglish("89");
        Student stu4=new Student();
        stu4.setNum("126");
        stu4.setName("张泰宁");
        stu4.setClassLevel("二年级");
        stu4.setMath("56");
        stu4.setChinese("67");
        stu4.setEnglish("67");
        Student stu5=new Student();
        stu5.setNum("127");
        stu5.setName("张泰宁");
        stu5.setClassLevel("三年级");
        stu5.setMath("50");
        stu5.setChinese("67");
        stu5.setEnglish("68");
        studentLists.add(stu1);
        studentLists.add(stu2);
        studentLists.add(stu4);
        studentLists.add(stu3);
        studentLists.add(stu5);
    }
    public static CellStyle generateDefaultCellStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        // 小五字体
        font.setFontHeight((short) (10.5 * 20));
        font.setFontName("宋体");
        style.setFont(font);
        // 水平/垂直对齐
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 自动换行
        style.setWrapText(true);
        // 单元格框线
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        return style;
    }
}
