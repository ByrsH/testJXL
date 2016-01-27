/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package TWorkbook;

/**
 *
 * @author Administrator
 */

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import jxl.Workbook;
import jxl.Sheet;
import jxl.Cell;
import jxl.Range;
import jxl.read.biff.BiffException;

public class Tworkbook {
    public static void main(String args[]) {
        try {
            File file = new File("D://JEtest/测试.xls");
            System.out.println("111");
            Workbook wb = Workbook.getWorkbook(file);
            System.out.println("111");
            //判断表是否受到保护   ?????文件加密进行getWorkbook() 时会抛出异常？？？？？？？？？？？？
            System.out.println("Is the sheet protect?   result: " + wb.isProtected());
            
            //在Excel表中可以选定多个单元格，给他们起一个名字，这里的range就代表的是一个你选定命名的那些单元格。
            //！！！！！！！！！！！！！！！！
            //注意不要用wps 操作选定一部分单元格给其命名，这样jxl用该函数会出现错误，返回的结果不正确。
            //！！！！！！！！！！！！！！！！
            //该方法返回所有你命名range的名字的字符串数组
            String [] rangeNameList = wb.getRangeNames();
            System.out.println("the range[]: " + Arrays.toString(rangeNameList));
            
            //有参数名，得到range[] 数组。    ？？？？range的名字是唯一的，为什么返回是数组呢？？？？
            Range[]  ranges;
            ranges = wb.findByName("name");
            System.out.println("the ranges[] length: " + ranges.length);   
                       
            
            //获取指定单元格， 参数格式为"sheet1!A1"
            System.out.println("111");
            Cell cellA1 = wb.getCell("第一页!A3");
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
            Sheet [] sheetlist;
            sheetlist = wb.getSheets();
            System.out.println("工作表个数为： " + sheetlist.length);
            
            //得到工作簿中所有的工作表名称，以字符串数组存储。
            String [] sheetnamelist = wb.getSheetNames();
            System.out.println("工作表个数：" + sheetnamelist.length + " 工作表名称： " + Arrays.toString(sheetnamelist));
            
            //得到jxl版本号
            String version = wb.getVersion();
            System.out.println(version);
            
            wb.close();
            
        }catch (IOException | BiffException | IndexOutOfBoundsException e) {
            System.out.println("Exception: " + e);
        }
    }
    
}
