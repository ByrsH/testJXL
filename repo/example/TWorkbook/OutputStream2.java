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

public class OutputStream2 {
    public static void main(String args[]) {
        try {

            //OutputStream os = new FileOutputStream("D://JEtest/测试4.xls",true);
            //创建一个可写的工作簿。当关闭工作簿时,它将直接流到输出流。通过这种方式,一个生成excel电子表格可以通过HTTP从servlet传递到浏览器            
                File file = new File("D://JEtest/测试2.xls"); 
               // OutputStream os = new FileOutputStream("D://JEtest/new.xls");
                File os = new File("D://JEtest/new4.xls");
                //OutputStream os = new FileOutputStream("D://JEtest/new2.xls");
                Workbook book = Workbook.getWorkbook(file);
                WritableWorkbook wb = Workbook.createWorkbook(os,book);
                //Workbook.createWorkbook(os);
                System.out.println("测试: " + wb.getNumberOfSheets());
                wb.write();
                wb.close();
              //  os.close();

            
        }catch (Exception e) {
            System.out.println("Exception: " + e);
        }
    }
    
}
