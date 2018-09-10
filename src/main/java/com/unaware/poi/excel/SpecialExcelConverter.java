package com.unaware.poi.excel;

import com.unaware.poi.excel.util.DataUtil;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

public class SpecialExcelConverter {

    /**
     * 根据html文件地址生成对应的xls
     * String title = doc.title();
     * 得到样式，以后可以根据正则表达式解析css，暂且没有找到cssparse
     * Elements style = doc.getElementsByTag("style");
     * <p>
     * 对于网页源码中的每一个table标签解析成一张表，追加到一个文件中
     * for (Element table : tables) {
     * //得到所有行
     * Elements trs = table.getElementsByTag("tr");
     * //得到列宽集合
     * Elements colgroups = table.getElementsByTag("colgroup");
     * }
     *
     * @param file 源文件地址
     * @throws IOException 输入输出异常
     */
    public static void toExcel(File file) throws IOException, WriteException {
        // 编码必须设置为 null,自动去识别编码,识别不出来默认为 UTF-8
        Document doc = Jsoup.parse(file, null, "");
        //得到Table
        Elements tables = doc.getElementsByTag("TABLE");
        if (tables.size() == 0) {
            return;
        }
        //得到所有行
        Elements trs = doc.getElementsByTag("tr");
        //得到列宽集合
        Elements colgroups = doc.getElementsByTag("colgroup");
        File file1 = new File("src\\test\\resources\\output\\" + DataUtil.getUUID() + ".xls");
        WritableWorkbook book = Workbook.createWorkbook(file1);
        WritableSheet sheet = book.createSheet("sheet1", 0);
        setColWidth(colgroups, sheet);
        mergeColRow(trs, sheet);
        book.write();
        book.close();
        Files.copy(file1.toPath(), file.toPath(), StandardCopyOption.REPLACE_EXISTING);
        Files.deleteIfExists(file1.toPath());
    }

    /**
     * 这个方法用于根据trs行数和sheet画出整个表格
     *
     * @param trs   源文件中的tr集合
     * @param sheet 工作表sheet
     * @throws RowsExceededException exception thrown when attempting to add a row to a spreadsheet which has already reached the maximum amount
     * @throws WriteException        写入异常
     */
    private static void mergeColRow(Elements trs, WritableSheet sheet) throws WriteException {
        int[][] rowhb = new int[300][50];
        for (int i = 0; i < trs.size(); i++) {
            Element tr = trs.get(i);
            Elements tds = tr.getElementsByTag("td");

            int realColNum = 0;
            for (Element td : tds) {
                if (rowhb[i][realColNum] != 0) {
                    realColNum = getRealColNum(rowhb, i, realColNum);
                }
                int rowSpan = 1, colSpan = 1;
                if (!"".equals(td.attr("rowspan"))) {
                    rowSpan = Integer.parseInt(td.attr("rowspan"));
                }
                if (!"".equals(td.attr("colspan"))) {
                    colSpan = Integer.parseInt(td.attr("colspan"));
                }
                String text = td.wholeText();
                drawMergeCell(rowSpan, colSpan, sheet, realColNum, i, text, rowhb);
                realColNum = realColNum + colSpan;
            }

        }
    }

    /**
     * 这个方法用于根据样式画出单元格，并且根据rowpan和colspan合并单元格
     *
     * @param rowspan    行合并单元格
     * @param colspan    列合并单元格
     * @param sheet      当前工作表sheet
     * @param realColNum 真实列号
     * @param realRowNum 真实行号
     * @param text       文本信息
     * @param rowhb      计数二维数组
     * @throws WriteException 写入异常
     */
    private static void drawMergeCell(int rowspan, int colspan, WritableSheet sheet, int realColNum, int realRowNum, String text, int[][] rowhb) throws WriteException {
        for (int i = 0; i < rowspan; i++) {
            for (int j = 0; j < colspan; j++) {
                if (i != 0 || j != 0) {
                    text = "";
                }
                Label label = new Label(realColNum + j, realRowNum + i, text);
                //设置单元格内容，字号12
                WritableFont countents = new WritableFont(WritableFont.TIMES, 10);
                WritableCellFormat cellf = new WritableCellFormat(countents);
                //把水平对齐方式指定为居中
                cellf.setAlignment(Alignment.LEFT);
                //把垂直对齐方式指定为居
                cellf.setVerticalAlignment(VerticalAlignment.CENTRE);
                label.setCellFormat(cellf);
                sheet.addCell(label);
                rowhb[realRowNum + i][realColNum + j] = 1;
            }
        }
        sheet.mergeCells(realColNum, realRowNum, realColNum + colspan - 1, realRowNum + rowspan - 1);
    }

    private static int getRealColNum(int[][] rowhb, int i, int realColNum) {
        while (rowhb[i][realColNum] != 0) {
            realColNum++;
        }
        return realColNum;
    }

    /**
     * 根据colgroups设置表格的列宽
     *
     * @param colgroups colgroup标签集合
     * @param sheet     当前工作表sheet
     */
    private static void setColWidth(Elements colgroups, WritableSheet sheet) {
        if (colgroups.size() > 0) {
            Element colgroup = colgroups.get(0);
            Elements cols = colgroup.getElementsByTag("col");
            for (int i = 0; i < cols.size(); i++) {
                Element col = cols.get(i);
                String strwd = col.attr("width");
                if (!"".equals(col.attr("width"))) {
                    int wd = Integer.parseInt(strwd);
                    sheet.setColumnView(i, wd / 8);
                }

            }

        }
    }
}
