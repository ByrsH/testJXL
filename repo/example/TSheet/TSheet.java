/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package TSheet;

/**
 *
 * @author Administrator
 */

import java.io.*;
import java.util.regex.Pattern;
import jxl.*;


public class TSheet {
    public static void main(String args[])  {
        try {
            File file = new File("D://JEtest/测试.xls");
            
            Workbook book = Workbook.getWorkbook(file);
            Sheet sheet = book.getSheet(0);
            
            //返回指定行列组成的单元格，如果指定的行列组合的单元格是合并后单元组的元素，将返回一个空的单元格，除非这个单元格是单元格组的第一个。
            //第一个参数是列，第二个参数是行。 索引从0开始
            Cell cell = sheet.getCell(0, 0);
            System.out.println("cell(0,0): " + cell.getContents());
            
            //返回指定位置的单元格，例如：“A4”。注意,这个方法与调用
            //getCell(CellReferenceHelper.getColumn(loc),CellReferenceHelper.getRow(loc))是相同的，
            //其隐含的性能开销为字符串解析。因此,这种方法应该谨慎使用。
            Cell cellh = sheet.getCell(CellReferenceHelper.getColumn("B1"),CellReferenceHelper.getRow("B1"));
            Cell cell2 = sheet.getCell("B1");      //索引从1开始
            System.out.println("cell(1,0): " + cell2.getContents());
            
            //返回工作表的行数
            int rows = sheet.getRows();
            System.out.println("表1的数据行数： " + rows);
            
            //返回工作表的列数
            int columns = sheet.getColumns();
            System.out.println("表1的数据列数： " + columns);
            
            //返回指定行的单元格数组
            Cell [] cellrowlist = sheet.getRow(0);
            System.out.println("表1第一行单元格个数： " + cellrowlist.length);
            
            //返回指定列的单元格数组
            Cell [] cellcolumnlist = sheet.getColumn(0);
            System.out.println("表1第一列单元格个数： " + cellcolumnlist.length);
            
            //返回工作表的名称
            String sheetname = sheet.getName();
            System.out.println("表1的名称： " + sheetname);
            
            //判断表是否是隐藏的            
            boolean judge = sheet.isHidden();
            System.out.println("表1是否被隐藏： " + judge);
            
            //判断表是否受保护
            boolean judeg2 = sheet.isProtected();
            System.out.println("表1是否受保护： " + judeg2);
            
            //得到内容与传进来字符串匹配的单元格。如果没有发现匹配，将返回null.
            //搜索执行是从最低行开始一行一行执行的，因此，行号越低，算法执行越高效。
            //返回第一个匹配的单元格。
            Cell cell3 = sheet.findCell("789.123");
            System.out.println("表1内容为test单元格的行数： " + cell3.getRow());
            System.out.println("表1内容为test的单元格内容： " + cell3.getContents());
            
            //得到内容与传进来字符串匹配的单元格。如果没有发现匹配，将返回null.
            //搜索执行是从最低行开始一行一行执行的，因此，行号越低，算法执行越高效。
            //Cell findCell(java.lang.String contents,int firstCol,int firstRow,int lastCol,int lastRow,boolean reverse)
            //contents是匹配内容，firstCol是搜索开始列，firstRow是搜索开始行，同样lastCol,lastRow是结束列和行
            //reverse 是指示是否执行反向搜索,反向搜索是从行、列数大的开始向小的方向搜索。true是执行反向搜索。
            //也是返回第一个匹配单元格。
            Cell cell4 = sheet.findCell("test",0,0,1,1,true);
            System.out.println("表1内容为test单元格的行数： " + cell4.getRow());
            
            //该函数后四个参数和上面函数一样，第一个参数是一个正则表达式。
            Pattern p = Pattern.compile("a*b");
            Cell cell5 = sheet.findCell(p, 0, 0, 2, 2, false);
            System.out.println("表1内容为正则表达式 a*b 单元格的行数： " + cell5.getRow());
            System.out.println("表1内容为正则表达式 a*b 单元格的内容是： " + cell5.getContents());
            
            //该函数与findCell前面的特性一样，这种方法与findCell方法的不同之处在于,只有标签单元格查询，
            //所有的数值单元格被忽略。这应该提高性能。
            LabelCell labelCell = sheet.findLabelCell("test");
            System.out.println("标签单元格内容： " + labelCell.getContents());
            
            //得到工作表中的超链接数组
            Hyperlink [] hyperlinkList = sheet.getHyperlinks();
            System.out.println("工作表中超链接个数： " + hyperlinkList.length);
            
            //得到工作表中合并单元格的range数组
            Range [] rangeList = sheet.getMergedCells();
            System.out.println("工作表中有多少个合并的单元格： " + rangeList.length);
            
            //得到用于表的settings
            SheetSettings sheetSet = sheet.getSettings();
            
            //得到指定列的列格式，如果没有特殊的格式则返回NULL
            jxl.format.CellFormat cellFormat = sheet.getColumnFormat(0);
            
            //返回指定列的宽度，如果没有特殊格式怎返回默认值
            int columnWidth = sheet.getColumnWidth(0);
            System.out.println("工作表第一列宽度： " + columnWidth);
            
            //得到指定列格式
            CellView cellViewCol = sheet.getColumnView(1);
            
            //得到指定行的高度         
            int rowHeight = sheet.getRowHeight(0);
            System.out.println("工作表第一行高度： " + rowHeight);
            
            //得到指定行格式
            CellView cellViewRow = sheet.getRowView(0);
            cellViewRow.setHidden(true);
            
            //得到工作表中图像数
            int numberOfImages = sheet.getNumberOfImages();
            System.out.println("工作表中图片个数： " + numberOfImages);
            
            //返回指定位置的图像,索引从0开始
            Image image = sheet.getDrawing(0);
            System.out.println("工作表中第一个图像的列位置： " + image.getColumn());

            //返回工作表的分页符
            int [] rowPageBreaks = sheet.getRowPageBreaks();
            //System.out.println("工作表分页符： " + rowPageBreaks[0]);
              
            //返回工作表的分页符
            int [] columnPageBreaks = sheet.getColumnPageBreaks();
            //System.out.println("工作表分页符： " + columnPageBreaks[0]);
            
            
            book.close();
        }catch (Exception e)  {
            System.out.println("Exception: " + e);
        }
    }
    
}
