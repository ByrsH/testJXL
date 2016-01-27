/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tsheetsettings;

/**
 *
 * @author Administrator
 */

import java.io.File;
import jxl.Cell;
import jxl.Workbook;
import jxl.Sheet;
import jxl.SheetSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class CreateLockSheet {
    public static void main(String [] args) {
        try {
            File file = new File("D://JEtest/测试.xls");
            
            Workbook book = Workbook.getWorkbook(file);
            WritableWorkbook wb = Workbook.createWorkbook(file, book);
            
            WritableSheet wSheet = wb.createSheet("name", 3);
            SheetSettings sheetset = wSheet.getSettings();
            sheetset.setPassword("123");
            sheetset.setProtected(true);
            wSheet.addCell(new Label(0,0,"yrs"));
            wSheet.addCell(new Label(0,1,"abc"));
            System.out.println(sheetset.getPassword());
             
            wb.write();
            wb.close();  
            book.close();
        }catch (Exception e) {
            System.out.println("Exception： " + e);
        }
    }
    
}
