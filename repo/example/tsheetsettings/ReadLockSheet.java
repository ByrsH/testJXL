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

public class ReadLockSheet {
    public static void main(String [] args) {
        try {
            File file = new File("D://JEtest/测试.xls");
            
            Workbook book = Workbook.getWorkbook(file);
            WritableWorkbook wb = Workbook.createWorkbook(file, book);
            
            WritableSheet wSheet = wb.getSheet("name");
            SheetSettings sheetset = wSheet.getSettings();

            wSheet.addCell(new Label(1,0,"yrs"));
            wSheet.addCell(new Label(1,1,"abc"));
            sheetset.setPassword("111"); 
            sheetset.setProtected(true);  
            System.out.println("表格密码：  " + sheetset.getPassword());
             
            wb.write();
            wb.close();  
            book.close();
        }catch (Exception e) {
            System.out.println("Exception： " + e);
        }
    }
    
}
