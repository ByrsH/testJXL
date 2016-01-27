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
import jxl.demo.Formulas;
import jxl.demo.XML;
import jxl.read.biff.BiffException;

/**
 *
 * @author Administrator
 */
public class TXML {
     public static void printXML(Workbook book,OutputStream out,String encoding,boolean f) throws IOException {
         //构造器
         //public XML(Workbook w,java.io.OutputStream out,java.lang.String encoding,booleam f)throws java.io.IOException
         //w - The workbook to interrogate
         //out - The output stream to which the XML values are written
        //enc - The encoding used by the output stream. Null or unrecognized values cause the encoding to default to UTF8
        //f - Indicates whether the generated XML document should contain the cell format information  
         //
        
        try {
            XML xml = new XML(book,out,encoding,f);
        } catch (IOException ex) {
            Logger.getLogger(TCSV2.class.getName()).log(Level.SEVERE, null, ex);
            throw ex;           
        }
    }
    
    public static void main(String [] args) throws Exception {
        try {
            File file = new File("D://JEtest/XML.xls");
            try (OutputStream out = new FileOutputStream("D://JEtest/printXML.XML")) {
                Workbook book = Workbook.getWorkbook(file);
                printXML(book,out,"utf8",true);
                book.close();
                out.close();
            }    
        }catch (IOException | BiffException e){
            System.out.println("Exception:  " + e);
            throw e;
        }
    }
    
}
