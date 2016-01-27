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
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class TImCopySheet {
    public static void main(String [] args) throws IOException, BiffException, WriteException  {
        try {
            Workbook book;
            WritableWorkbook wb;
           
            File file = new File("D://JEtest/copy.xls");

            if(file.exists()) {
                //如果文件存在

                book = Workbook.getWorkbook(file);
                wb = Workbook.createWorkbook(file,book);
                
                wb.copySheet(0, "new", 1);
                wb.copySheet("copy", "new2", 2); 
                
                
                wb.copySheet(0, "yrs", 0);
                wb.write();
                wb.close();
                book.close();
            }
            else {
                System.out.println("文件不存在");
            }
                
        }catch (IOException | BiffException | WriteException e) {
            System.out.println("Exception:  " + e);
            throw e;
        }
    }
    
}
