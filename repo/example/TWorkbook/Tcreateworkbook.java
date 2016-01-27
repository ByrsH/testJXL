/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package TWorkbook;

/**
 *
 * @author 杨汝生
 */

import java.io.*;
import jxl.*;
import jxl.write.*;

public class Tcreateworkbook {
    public static void main(String args[]) {
        try {
            File file = new File("D://JEtest/测试2.xls");
            //创建一个可写的工作簿
            if (!file.exists()) {   
                //如果文件不存在，创建可写工作簿的方式
                WritableWorkbook wb1 = Workbook.createWorkbook(file);           
                System.out.println("sheet numbers: " + wb1.getNumberOfSheets());
                WritableSheet wsheet1 = wb1.createSheet("sheet1", 0);
       
                System.out.println("sheet numbers: " + wb1.getNumberOfSheets());
                wb1.write();
                wb1.close();
            }
            else {
                //如果文件存在，创建可写工作簿方式
                //System.out.println("************************************a");
                Workbook book = Workbook.getWorkbook(file);
                //System.out.println("************************************b");
                WritableWorkbook wb1 = Workbook.createWorkbook(file, book);
                //System.out.println("************************************c");
                System.out.println("sheet numbers: " + wb1.getNumberOfSheets());
                wb1.write();      //注意一定要用write()函数写回去，不然会出错。
                wb1.close();
               
            }
            
            
            File file2 = new File("D://JEtest/测试3.xls");
            //WorkbookSettings工作簿的本地属性设置。
            if (!file2.exists()) {
                //如果文件不存在，创建可写工作簿
                
                //ws 没有设置相关的属性
                WorkbookSettings ws = new WorkbookSettings();
               
                WritableWorkbook wb2 = Workbook.createWorkbook(file2, ws);
                WritableSheet wsheet2 = wb2.createSheet("sheet2", 0);
                System.out.println("sheet2 numbers: " + wb2.getNumberOfSheets());
                wb2.write();
                wb2.close();
            }
            else {
                //如果文件存在，创建可写工作簿
                WorkbookSettings ws = new WorkbookSettings();
                Workbook book2 = Workbook.getWorkbook(file2);
                WritableWorkbook wb2 = Workbook.createWorkbook(file2, book2, ws);
                System.out.println("sheet2 numbers: " + wb2.getNumberOfSheets());
                wb2.write();
                wb2.close();
            }
            
            
            //OutputStream os = new FileOutputStream("D://JEtest/测试4.xls",true);
            //创建一个可写的工作簿。当关闭工作簿时,它将直接流到输出流。通过这种方式,一个生成excel电子表格可以通过HTTP从servlet传递到浏览器            
            File file3 = new File("D://JEtest/测试4.xls");
            if (!file3.exists()) {
                //如果文件不存在
                OutputStream os = new FileOutputStream("D://JEtest/测试4.xls");
                System.out.println("1111111111111 " );
                WritableWorkbook wb3 = Workbook.createWorkbook(os);
                WritableSheet wsheet3 = wb3.createSheet("sheet3", 0);
                System.out.println("sheet3 numbers: " + wb3.getNumberOfSheets());
                wb3.write();
                wb3.close();
                os.close();   
            }
            else {
                //有问题？？？？？？？？？？？？？？？？？？？？？？？
                ///*
                //如果文件存在
                OutputStream os = new FileOutputStream("D://JEtest/test.xls");
                System.out.println("222222222222 " );
                
                Workbook book3 = Workbook.getWorkbook(file3);
                System.out.println("33333333333 " );
                WritableWorkbook wb5 = Workbook.createWorkbook(os, book3);
                System.out.println("测试: " + wb5.getNumberOfSheets());
                WritableSheet wsheet3 = wb5.createSheet("测试4", 0);
                System.out.println("测试: " + wb5.getNumberOfSheets());
                 System.out.println("33333333ss333 " );
                //wb5.write();
                 System.out.println("33333333sssss333 " );
                wb5.close();
                //book3.close();
                 //System.out.println("33333333ss333 " );
                //os.write(1);
                os.close();
                //*/
            }
            //System.out.println("33333333ss333 " );
            
            //设置相关属性
            OutputStream os2 = new FileOutputStream("D://JEtest/测试6.xls");
            WorkbookSettings ws4 = new WorkbookSettings(); 
            WritableWorkbook wb4 = Workbook.createWorkbook(os2,ws4);
            WritableSheet wsheet4 = wb4.createSheet("sheet4", 0);
            System.out.println("sheet4 numbers: " + wb4.getNumberOfSheets());
            wb4.write();
            wb4.close();
            os2.close();
            
            //?????????????????????????
           //WritableWorkbook wb5 = Workbook.createWorkbook(os3,book4,ws5)
            
            
        }catch (Exception e) {
            System.out.println("Exception: " + e);
        }
    }
    
}
