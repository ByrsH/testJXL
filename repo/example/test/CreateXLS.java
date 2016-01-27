/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package test;

/**
 *
 * @author yrs
 */

import java.io.*;
import jxl.*;
import jxl.write.*;


public class CreateXLS {
    public static void main(String args[]) {
        try{
            WritableWorkbook book = Workbook.createWorkbook(new File("D:/JEtest/测试.xls"));  //工作簿
            WritableSheet sheet = book.createSheet("第一页", 0);             //工作表         
            Label label = new Label(0,0,"test");                    //单元格
            sheet.addCell(label);
            jxl.write.Number number = new jxl.write.Number(1, 0, 789.123);
            sheet.addCell(number);
            book.write();
            book.close();
        }catch(Exception e) {
            System.out.println("异常 "+ e);
        }
    }
}
