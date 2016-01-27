/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package test;


import java.io.File;
import java.io.IOException;
import jxl.Workbook;
import jxl.Sheet;
import jxl.Cell;
import jxl.read.biff.BiffException;

/**
 *
 * @author Administrator
 */
public class ReadXLS {
    public static void main(String args[]) {
        try{
            Workbook book = Workbook.getWorkbook(new File("D:/JEtest/测试.xls"));
            Sheet sheet = book.getSheet(0);
            Cell cell1 = sheet.getCell(0, 0);
            String result = cell1.getContents();
            //book.
            
            System.out.println(result);
            book.close();
        }catch(IOException | BiffException | IndexOutOfBoundsException e) {
            System.out.println(e);
        }
    }
    
}
