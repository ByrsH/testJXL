/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package twritableworkbook;

/**
 *
 * @author Administrator
 */

import java.io.File;
import java.io.IOException;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class TImportSheet {
    public static void main(String [] args) throws IOException, WriteException, BiffException  {
        //try {
            Workbook book;
            WritableWorkbook wb;    
            
            File fileTarget = new File("D://JEtest/target.xls");
            File fileSource = new File("D://JEtest/source.xls");

            book = Workbook.getWorkbook(fileSource);                
            System.out.println("SSSSSSSSSSSSSS");
            wb = Workbook.createWorkbook(fileTarget,book);
            System.out.println("SSSSSSSSSSSSSS");
            wb.importSheet("ss", 0, book.getSheet(0));
                
            book.close();
            wb.write();
            wb.close();
                    
        //}catch (IOException | BiffException | IndexOutOfBoundsException | WriteException e) {}
            //System.out.println("Exception:  " + e);
        
    }
    
}
