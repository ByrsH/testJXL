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
import jxl.*;
import jxl.write.*;

public class UpdateXLS {
    public static void main(String args[]) {
        try {
            Workbook wb = Workbook.getWorkbook(new File("D://JEtest/测试.xls"));     //Excel 获得文件
            //打开一个文件副本，并指定数据写回到原文件
            WritableWorkbook book = Workbook.createWorkbook(new File("D://JEtest/测试.xls"), wb);
      
            //打开一个文件，覆盖原来的内容
            //WritableWorkbook book = Workbook.createWorkbook(new File("D://JEtest/测试.xls"));
            WritableSheet sheet = book.createSheet("第二页", 1);
            sheet.addCell(new Label(0,0,"第二页的测试数据"));
            book.write();
            book.close();
            wb.close();
            
        }catch (Exception e){
            System.out.println(e);
        }
    }
    
}
