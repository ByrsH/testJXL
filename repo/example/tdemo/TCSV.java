/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tdemo;

/**
 *
 * @author Administrator
 */

import jxl.demo.CSV;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class TCSV {
    public static void main(String [] args) throws Exception {
        //public CSV(Workbook w,java.io.OutputStream out,java.lang.String encoding,boolean hide)throws java.io.IOException
        //w - The workbook to interrogate
        //out - The output stream to which the CSV values are written
        //encoding - The encoding used by the output stream. Null or unrecognized values cause the encoding to default to UTF8
        //hide - Suppresses hidden cells 

        try {
            File file = new File("D://JEtest/CSV.xls");
            try (OutputStream out = new FileOutputStream("D://JEtest/OutCSV.txt")) {
                Workbook book = Workbook.getWorkbook(file);
                
                CSV csv = new CSV(book,out,"utf8",false);
                
                book.close();
                out.close();
            }
            
            
        }catch (IOException | BiffException e){
            System.out.println("Exception:  " + e);
            throw e;
        }finally {
            
        }
        
    }
    
}
