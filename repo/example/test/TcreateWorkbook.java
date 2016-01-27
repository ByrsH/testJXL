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

import jxl.*;
import java.io.*;
import jxl.write.*;

public class TcreateWorkbook {
    public static void main(String args[]) {
        try {
            File file = new File("D://JEtest/测试.xls");
            OutputStream out = new FileOutputStream("D://JEtest/测试2.xls");
        
            WritableWorkbook wb1 = Workbook.createWorkbook(file);
            
            WritableWorkbook wb2 = Workbook.createWorkbook(out);
            
            
            wb2.write();
            wb1.close();
        }catch (Exception e) {
            System.out.println("Exception: " + e);
        }
        
    }
    
}
