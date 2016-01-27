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

public class Tworkbook {
    public static void main(String args[]) {
        try {
            Workbook wb = Workbook.getWorkbook(new File("D://JEtest/测试.xls"));
            
            //判断表是否受到保护
            System.out.println("Is the sheet protect?   result: " + wb.isProtected());
            
            //System.out.println("getRangeNames: " + Arrays.toString(wb.getRangeNames()));      //不知道什么用
            //System.out.println("the name range: " + wb.findByName("7699a589").toString());   //不知道什么用
            
            //获取指定单元格， 参数格式为"sheet1!A1"
            Cell cellA1 = wb.getCell("第一页!A1");
            System.out.println("第一页里A1的内容： " + cellA1.getContents());
            
            //返回工作簿中工作表个数
            int sheetnumber = wb.getNumberOfSheets();
            System.out.println("工作表个数为： " + sheetnumber);
            
            //得到指定的表，根据索引值，从零开始。  不要做不必要的调用，因为每次调用都要重读工作表， 
            //此外客户端也不要持有不必要的表的引用，这回阻止垃圾回收器释放内存。
            Sheet sheet0 = wb.getSheet(0);
            System.out.println("第一个表的名字是： " + sheet0.getName());
            //这种方式是根据工作表的名称来得到工作表的。
            Sheet sheet1 = wb.getSheet("第一页");
            
            //得到工作表数组
            Sheet [] sheetlist = wb.getSheets();
            
            //得到工作簿中所有的工作表名称，以字符串数组存储。
            String [] sheetnamelist = wb.getSheetNames();
            System.out.println("工作表个数：" + sheetnamelist.length + " 工作表名称： " + Arrays.toString(sheetnamelist));
            
            //得到jxl版本号
            String version = wb.getVersion();
            System.out.println(version);
            
            wb.close();
            
        }catch (Exception e) {
            System.out.println("Exception: " + e);
        }
    }
    
}
