package com.poi.ykk.utils;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.util.BytePictureUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Color;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

import static com.poi.ykk.utils.FileUtils.getUploadFileAbsPath;
import static com.poi.ykk.utils.FileUtils.isExist;

/**
 * 模板工具类
 */
public class TplUtils {

    private static final Logger logger = LoggerFactory.getLogger(TplUtils.class);

    // 静态变量方便调用其状态无关的方法
    private static final NiceXWPFDocument document = new NiceXWPFDocument();

    /**
     * 替换模板
     *
     * @param tplPath    模板路径
     * @param outputName 输出文件名
     * @param replaceMap 替换数据映射
     * @return true-成功 false-失败
     */
    public static boolean replaceTemplate(
        Path tplPath,
        String outputName,
        Map<String, Object> replaceMap
    ) {
        try {
            XWPFTemplate tpl = XWPFTemplate.compile(tplPath.toString())
                .render(replaceMap);
            FileOutputStream outputStream = new FileOutputStream(
                Paths.get("report", outputName).toString());
            tpl.write(outputStream);
            outputStream.flush();
            outputStream.close();
            tpl.close();
            return true;
        } catch (IOException e) {
            logger.warn(e.getMessage());
            return false;
        }
    }

    /**
     * 替换模板
     *
     * @param tpl        poi-tl 模板对象
     * @param outputName 输出文件名
     * @param replaceMap 替换数据映射
     * @return 成功与否
     */
    public static boolean replaceTemplate(
        XWPFTemplate tpl,
        String outputName,
        Map<String, Object> replaceMap
    ) {
        try {
            tpl.render(replaceMap);
            FileOutputStream outputStream = new FileOutputStream(
                Paths.get("report", outputName).toString());
            tpl.write(outputStream);
            outputStream.flush();
            outputStream.close();
            tpl.close();
            return true;
        } catch (Exception e) {
            logger.warn(e.getMessage());
            return false;
        }
    }

    /**
     * 缺少图片内容时填充一个占位图片
     *
     * @param width  宽度
     * @param height 高度
     */
    public static PictureRenderData generateImageHolder(int width, int height) {
        BufferedImage bufferImage = BytePictureUtils.newBufferImage(width, height);
        Graphics2D g = (Graphics2D) bufferImage.getGraphics();
        g.setColor(Color.WHITE);
        g.fillRect(0, 0, width, height);
        g.dispose();
        bufferImage.flush();
        return new PictureRenderData(width, height, ".png", BytePictureUtils.getBufferByteArray(bufferImage));
    }

    /**
     * 处理图片数据，当不存在时填充占位图片
     *
     * @param map    数据映射
     * @param key    key
     * @param width  宽度
     * @param height 高度
     */
    public static void handleImageField(
        Map<String, Object> map,
        String key,
        int width,
        int height
    ) {
        try {
            String photoFullName = getUploadFileAbsPath((String) map.get(key));
            if (isExist(photoFullName)) {
                map.put(key, new PictureRenderData(width, height, photoFullName));
            } else {
                map.put(key, generateImageHolder(width, height));
            }
        } catch (Exception e) {
            logger.warn(e.getMessage());
        }
    }
    /**
     * 清除表格行内的占位数据
     *
     * @param row
     */
    public static void clearTableRowParagraphRuns(XWPFTableRow row) {
        try {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    int size = paragraph.getRuns().size();
                    for (int i = 0; i < size; i++) {
                        paragraph.removeRun(i);
                    }
                }
            }
        } catch (Exception e) {
            logger.warn(e.getMessage());
        }
    }
    

    /**
     * 设置表格行水平/垂直居中对齐
     *
     * @param row XWPFTableRow
     */
    public static void setTableRowAlignCenter(XWPFTableRow row) {
        try {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                // 垂直居中
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                // 水平居中
                CTTc cttc = cell.getCTTc();
                CTP ctp = cttc.getPList().get(0);
                CTPPr ctppr = ctp.getPPr();
                if (ctppr == null) {
                    ctppr = ctp.addNewPPr();
                }
                CTJc ctjc = ctppr.getJc();
                if (ctjc == null) {
                    ctjc = ctppr.addNewJc();
                }
                ctjc.setVal(STJc.CENTER);
            }
        } catch (Exception e) {
            logger.warn(e.getMessage());
        }

    }

    /**
     * 垂直合并单元格
     *
     * @param table   表格对象
     * @param col     合并所在的列
     * @param fromRow 起始行数
     * @param toRow   结束行数
     */
    public static void mergeCellsVertically(
        XWPFTable table,
        int col,
        int fromRow,
        int toRow
    ) {
        try {
            document.mergeCellsVertically(table, col, fromRow, toRow);
        } catch (Exception e) {
            logger.warn(e.getMessage());
        }
    }
    /**
     * 养护维修垂直合并单元格
     *
     * @param table   表格对象
     * @param col     合并所在的列
     * @param row   用于合并的行数集合
     */
    public static void maintainMergeCellsVertically(
        XWPFTable table,
        int col,
        List<Integer> row
    ) {
        try {
            for(int i=0;i<row.size();i++) {
                int fromRow=row.get(i);
                int toRow=row.get(i+1);
                document.mergeCellsVertically(table, col, fromRow, toRow);
                i++;
            }
        } catch (Exception e) {
            logger.warn(e.getMessage());
        }
    }

    /**
     * 垂直合并单元格
     *
     * @param table         表格对象
     * @param col           合并所在的列
     * @param mergeRowCount 间隔行标数组
     */
    public static void mergeCellsVertically(
        XWPFTable table,
        int col,
        List<Integer> mergeRowCount
    ) {
        if (table == null || col < 0 || mergeRowCount == null) {
            return;
        }
        for (int i = 0, size = mergeRowCount.size(); i < size - 1; i++) {
            int fromRow = mergeRowCount.get(i);
            if (i + 1 < size) {
                int toRow = mergeRowCount.get(i + 1) - 1;
                if (fromRow < toRow) {
                    mergeCellsVertically(table, col, fromRow, toRow);
                }
            }
        }
    }

    /**
     * 水平合并单元格
     *
     * @param table   表格对象
     * @param row     行
     * @param fromCol 起始列
     * @param toCol   结束列
     */
    public static void mergeCellsHorizontally(
        XWPFTable table,
        int row,
        int fromCol,
        int toCol
    ) {
        document.mergeCellsHorizonal(table, row, fromCol, toCol);
    }

    /**
     * 新行创建新列
     *
     * @param row     行对象
     * @param colSize 列数
     */
    public static void addCells(XWPFTableRow row, int colSize) {
        if (row != null) {
            for (int j = row.getTableCells().size(); j < colSize; j++)
                row.addNewTableCell();
        }
    }

    /**
     * 默认表格添加文本样式
     *
     * @param cell 表格单元
     * @param text 文本内容
     */
    public static void addRun(XWPFTableCell cell, String text) {
        if (cell == null) {
            return;
        }
        XWPFParagraph paragraph;
        try {
            paragraph = cell.getParagraphs().get(0);
        } catch (Exception e) {
            logger.warn(e.getMessage());
            paragraph = cell.addParagraph();
        }
        addRun(paragraph.createRun(),
            "宋体", 9, "000000", text,
            false, false, true
        );
    }

    /**
     * 添加文本
     * link: https://stackoverflow.com/questions/27634991/how-to-format-the-text-in-a-xwpftable-in-apache-poi/29258785?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
     *
     * @param run        文本对象
     * @param fontFamily 字体类型
     * @param fontSize   字体大小
     * @param colorRGB   字体颜色
     * @param text       文本内容
     * @param bold       是否加粗
     * @param addBreak   是否换行
     */
    public static void addRun(
        XWPFRun run,
        String fontFamily,
        int fontSize,
        String colorRGB,
        String text,
        boolean bold,
        boolean addBreak,
        boolean isAsciiTimesNewRoman
    ) {

        if (run == null) {
            return;
        }
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        run.setColor(colorRGB);
        run.setText(text);
        run.setBold(bold);
        if (addBreak) {
            run.addBreak();
        }
        // 对于英文数字等使用新罗马字体
        if (isAsciiTimesNewRoman) {
            run.setFontFamily("Times New Roman", XWPFRun.FontCharRange.ascii);
        }
    }

    /**
     * 清除表格框线
     *
     * @param row 表格行
     */
    public static void clearTableBorders(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            try {
                CTTc ctTc = cell.getCTTc();
                CTTcPr tcPr = ctTc.getTcPr();
                if (tcPr == null) {
                    tcPr = ctTc.addNewTcPr();
                }
                CTTcBorders borders = tcPr.getTcBorders();
                if (borders == null) {
                    borders = tcPr.addNewTcBorders();
                }
                borders.addNewRight().setVal(STBorder.NIL);
                borders.addNewLeft().setVal(STBorder.NIL);
                borders.addNewBottom().setVal(STBorder.NIL);
                borders.addNewTop().setVal(STBorder.NIL);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 复制表格
     * link: https://stackoverflow.com/questions/48322534/apache-poi-how-to-copy-tables-from-one-docx-to-another-docx?rq=1&utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
     *
     * @param source 模板
     * @param target 目标
     */
    public static void copyTable(XWPFTable source, XWPFTable target) {
        target.getCTTbl().setTblPr(source.getCTTbl().getTblPr());
        target.getCTTbl().setTblGrid(source.getCTTbl().getTblGrid());
        for (int r = 0; r < source.getRows().size(); r++) {
            XWPFTableRow targetRow = target.createRow();
            XWPFTableRow row = source.getRows().get(r);
            targetRow.getCtRow().setTrPr(row.getCtRow().getTrPr());
            for (int c = 0; c < row.getTableCells().size(); c++) {
                // newly created row has 1 cell
                XWPFTableCell targetCell = c == 0 ? targetRow.getTableCells().get(0) : targetRow.createCell();
                XWPFTableCell cell = row.getTableCells().get(c);
                targetCell.getCTTc().setTcPr(cell.getCTTc().getTcPr());
                XmlCursor cursor = targetCell.getParagraphArray(0).getCTP().newCursor();
                for (int p = 0; p < cell.getBodyElements().size(); p++) {
                    IBodyElement elem = cell.getBodyElements().get(p);
                    if (elem instanceof XWPFParagraph) {
                        XWPFParagraph targetPar = targetCell.insertNewParagraph(cursor);
                        cursor.toNextToken();
                        XWPFParagraph par = (XWPFParagraph) elem;
                        copyParagraph(par, targetPar);
                    } else if (elem instanceof XWPFTable) {
                        XWPFTable targetTable = targetCell.insertNewTbl(cursor);
                        XWPFTable table = (XWPFTable) elem;
                        copyTable(table, targetTable);
                        cursor.toNextToken();
                    }
                }
                // newly created cell has one default paragraph we need to remove
                targetCell.removeParagraph(targetCell.getParagraphs().size() - 1);
            }
        }
        // newly created table has one row by default. we need to remove the default row.
        target.removeRow(0);
    }

    /**
     * 复制段落
     *
     * @param source 模板
     * @param target 目标
     */
    public static void copyParagraph(XWPFParagraph source, XWPFParagraph target) {
        target.getCTP().setPPr(source.getCTP().getPPr());
        for (int i = 0; i < source.getRuns().size(); i++) {
            XWPFRun run = source.getRuns().get(i);
            XWPFRun targetRun = target.createRun();
            // copy formatting
            targetRun.getCTR().setRPr(run.getCTR().getRPr());
            // no images just copy text
            targetRun.setText(run.getText(0));
        }
    }

    /**
     * 填充行
     *
     * @param row  行对象
     * @param data 数据列
     */
    public static void fillTableRow(XSSFRow row, String[] data, CellStyle style) {
        for (int i = 0, size = data.length; i < size; i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(data[i]);
            cell.setCellStyle(style);
            cell.setCellType(CellType.STRING);
        }
    }

    /**
     * 默认单元格样式
     * link: https://stackoverflow.com/questions/28539726/how-to-set-font-in-decimal-number-using-hssffont?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
     *
     * @return 样式
     */
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


    /**
     * 合并单元格
     *
     * @param sheet      表页
     * @param mergeCount 合并位置
     * @param index      合并行/列
     * @param isVertical 是否垂直合并
     */
    public static void mergeCell(
        XSSFSheet sheet,
        List<Integer> mergeCount,
        int index,
        boolean isVertical
    ) {
        for (int i = 0, size = mergeCount.size(); i < size; i++) {
            int from = mergeCount.get(i);
            int to = i + 1 < size ? mergeCount.get(i + 1) - 1 : -1;
            if (from < to) {
                if (isVertical) {
                    sheet.addMergedRegion(new CellRangeAddress(from, to, index, index));
                } else {
                    sheet.addMergedRegion(new CellRangeAddress(index, index, from, to));
                }
            }
        }
    }
}
