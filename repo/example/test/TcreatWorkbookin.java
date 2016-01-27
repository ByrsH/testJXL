package test;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Administrator
 */

import jxl.*;
import java.io.*;
import java.util.Arrays;
import jxl.read.biff.BiffException;
import jxl.write.*;

public class TcreatWorkbookin {
    public static void main(String args[]) throws FileNotFoundException, IOException, BiffException, WriteException {
        //File file = new File("D://JEtest/测试.xls");
        InputStream is = new FileInputStream("D://JEtest/测试.xls");
        OutputStream os = new FileOutputStream("D://JEtest/test.xls");
        
        Workbook book = Workbook.getWorkbook(is);
        WritableWorkbook wb = Workbook.createWorkbook(os,book);
//        System.out.println(wb.getSheet(0).getCell(0,0).getContents());
//        WritableSheet wsheet = wb.createSheet("添加", 1);//????????????????????????
//        Label labelCF = new Label(0, 0, "hello");// 创建写入位置和内容 
//        wsheet.addCell(labelCF);// 将Label写入sheet中 
//        System.out.println(wb.getSheet(0).getCell(0,0).getContents());
        wb.write();
        wb.close();
        os.close();
    }
    
}
