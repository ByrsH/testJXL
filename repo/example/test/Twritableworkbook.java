/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package test;

/**
 *
 * @author Administrator
 */

import java.io.*;
import java.util.Arrays;
import jxl.*;
import jxl.write.*;

public class Twritableworkbook {
    public static void main(String args[]) {
        try {
            Workbook wb = Workbook.getWorkbook(new File("D://JEtest/测试.xls"));
            
            //打开一个文件副本，并且指定数据写回到原文件
            WritableWorkbook book = Workbook.createWorkbook(new File("D://JEtest/测试.xls"),wb);
            
            //得到 WritableSheet 数组。
            WritableSheet [] wsheetlist = book.getSheets();
            System.out.println("工作表个数： " + wsheetlist.length);
            
            //得到工作表名称字符串数组
            
            String [] sheetnamelist = book.getSheetNames();
            System.out.println("工作表名称： " + Arrays.toString(sheetnamelist));
            
            //得到指定的表，根据索引值，从零开始。  不要做不必要的调用，因为每次调用都要重读工作表， 
            //此外客户端也不要持有不必要的表的引用，这回阻止垃圾回收器释放内存。
            Sheet sheet0 = book.getSheet(0);
            System.out.println("第一个表的名字是： " + sheet0.getName());
            //这种方式是根据工作表的名称来得到工作表的。
            Sheet sheet1 = book.getSheet("第一页");
            
            //获取指定单元格， 参数格式为"sheet1!A1"
            Cell cellA1 = book.getWritableCell("第一页!A1");
            System.out.println("第一页里A1的内容： " + cellA1.getContents());
            
            //返回工作簿中工作表个数
            int sheetnumber = book.getNumberOfSheets();
            System.out.println("工作表个数为： " + sheetnumber);
            
            
            //创建一个WritableSheet,第一个参数为表名，第二个参数为表的索引。 如果索引值小于等于0，怎在工作簿开始处创建，
            //如果值大于表数，则在最后创建表。
            WritableSheet wsheet = book.createSheet("第三页", 3);
            book.write();
            
            /*   不知道怎么使用？？？？？？
            //importSheet 是导入一个不同工作簿的表，所有元素进行深拷贝。
            Workbook wb2 = Workbook.getWorkbook(new File("D://JEtest/write.xls"));
            Sheet sheet = wb2.getSheet(0);
            WritableSheet wsheet2 = book.importSheet("第四页", 0, sheet);
            wb2.close();
            */
            
            
            book.copySheet(0, "复制", 2);
            
            
            book.write();
            book.close();
            wb.close();
        }catch (Exception e) {
            System.out.println("Exception: " + e);
        }
    }
    
}
