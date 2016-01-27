/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tdemo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.Workbook;
import jxl.demo.EscherDrawingGroup;
import jxl.read.biff.BiffException;

/**
 *
 * @author Administrator
 */
public class TEscherDrawingGroup {
     public static void displayEscherDrawingGroup(Workbook book,OutputStream out,String encoding) throws IOException {
         //构造器
         //public EscherDrawingGroup(Workbook w,java.io.OutputStream out,java.lang.String encoding)throws java.io.IOException
         //w - The workbook to interrogate
         //out - The output stream to which the CSV values are written
         //encoding - The encoding used by the output stream. Null or unrecognized values cause the encoding to default to UTF8 
         //输出的结果是各个工作表的名称。

        
        try {
            EscherDrawingGroup escher = new EscherDrawingGroup(book,out,encoding);
        } catch (IOException ex) {
            Logger.getLogger(TCSV2.class.getName()).log(Level.SEVERE, null, ex);
            throw ex;           
        }
    }
    
    public static void main(String [] args) throws Exception {
        try {
            File file = new File("D://JEtest/EscherDrawingGroup.xls");
            try (OutputStream out = new FileOutputStream("D://JEtest/displayEscherDrawingGroup.txt")) {
                Workbook book = Workbook.getWorkbook(file);
                displayEscherDrawingGroup(book,out,"utf8");
                book.close();
                out.close();
            }    
        }catch (IOException | BiffException e){
            System.out.println("Exception:  " + e);
            throw e;
        }
    }
    
}
