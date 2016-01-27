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

import java.io.*;
import jxl.*;

public class Tgetworkbook {
    public static void main(String args[])  {
        try {
            //文件已经存在，得到工作簿、
            File file = new File("D://JEtest/测试.xls");
            Workbook book1 = Workbook.getWorkbook(file);
            System.out.println("sheet's number: " + book1.getNumberOfSheets());
            book1.close();
            
            WorkbookSettings ws = new WorkbookSettings();
            Workbook book2 = Workbook.getWorkbook(file,ws);
            System.out.println("sheet's number: " + book2.getNumberOfSheets());
            book2.close();
            
            InputStream is = new FileInputStream(file);
            Workbook book3 = Workbook.getWorkbook(is);
            System.out.println("sheet's number: " + book3.getNumberOfSheets());
            book3.close();
            is.close();
            
            InputStream is2 = new FileInputStream(file);
            Workbook book4 = Workbook.getWorkbook(is2, ws);
            System.out.println("sheet's number: " + book4.getNumberOfSheets());
            book4.close();
            is2.close();
            
            
        }catch (Exception e) {
            System.out.println("Exception: " + e);
        }
    }
    
}
