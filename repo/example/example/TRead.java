/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package example;

import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 *
 * @author Administrator
 */
public class TRead {
    
    public static void read (Workbook book) {
        int rows,cols;
        int sheetNumber = book.getNumberOfSheets();
        String [] sheetNameList = book.getSheetNames();
        Sheet [] sheetList = book.getSheets();
        
        for(int i = 0;i < sheetNumber;i++) {
            System.out.println("##############" + sheetNameList[i] + "##############");
            rows = sheetList[i].getRows();
            for(int j = 0;j < rows;j++) {
                Cell [] cellList = sheetList[i].getRow(j);
                for (Cell cell : cellList) {
                    System.out.print(cell.getContents() + "  ");
                }
                System.out.println();
            }          
        }      
    }
    public static void main(String [] args) throws IOException, BiffException  {
        try {
            File file = new File("D://JEtest//read.xls");
            
            Workbook book = Workbook.getWorkbook(file); 
            TRead.read(book);
            book.close();            
        }catch (IOException | BiffException e) {
            System.out.println("Exception:  " + e);
            throw e;
        }
    }  
}
