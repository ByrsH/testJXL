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
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class TImportSheet2 {
    public static void main(String [] args)  {
        File sourceFile;
        File targetFile;
        sourceFile = new File("D://JEtest/gaos.xls");
        targetFile = new File("D://JEtest/importFile4.xls");
        try {
            Workbook wb;
            WritableWorkbook newWb;
            try (InputStream fis = new FileInputStream(sourceFile)) {
                wb = Workbook.getWorkbook(fis);
                newWb = Workbook.createWorkbook(targetFile, wb);
                newWb.importSheet("NewSheet", 0, wb.getSheet(0));
            }
            wb.close();
            newWb.write();
            newWb.close();
                    
        }catch (IOException | BiffException | IndexOutOfBoundsException | WriteException e){}
            //System.out.println("Exception:  " + e);
        
    }
    
}
