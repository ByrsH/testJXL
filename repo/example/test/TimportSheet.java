/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package test;

/**
 *
 * @author Administrator
 */


import java.io.File;
import java.io.IOException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;



public class TimportSheet {
    public static void main(String args[]) {
        try {
            
            Workbook from_book = Workbook.getWorkbook(new File("D://JEtest/write.xls"));
            //Workbook to_book = Workbook.getWorkbook(new File("D://JEtest/测试.xls"));
            WritableWorkbook wb = Workbook.createWorkbook(new File("D://JEtest/测试.xls"));
            //WritableWorkbook wb = Workbook.createWorkbook(new File("D://JEtest/测试.xls"), to_book);
            
            Sheet from_sheet = from_book.getSheet(0);
            System.out.println("from_book sheet[0] name: " + from_sheet.getName());
            
            WritableSheet wsheet = wb.importSheet("aa", 0, from_sheet);
            System.out.println(wsheet.getName());
            
            wb.write();
            from_book.close();
            wb.close();
           

        }catch (Exception e) {
            System.out.println("Exception: " + e);
        }
    }
	  
} 
	