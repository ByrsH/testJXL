/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tdemo;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.demo.ReadWrite;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

/**
 *
 * @author Administrator
 */
public class TReadWrite {
     public static void copyReadWrite(String input,String output) throws IOException, BiffException, WriteException {
        //构造器
        //ReadWrite(java.lang.String input, java.lang.String output) 
         
        //public void readWrite()throws java.io.IOException,jxl.read.biff.BiffException,WriteException
        //读一个Excel文件，然后克隆一个新文件。input 是源文件，output是产生的新文件。
        //如果读进的表格是被叫做 jxlrwtest.xls(提供分布），那么这个类将改变某些字段在复制的表格中。
        //这说明它是可能的：读一个电子表格，改变一些值，写进一个新的文件中。
        
        try {
            ReadWrite copyRW = new ReadWrite(input,output);
            copyRW.readWrite();
        } catch (IOException ex) {
            Logger.getLogger(TCSV2.class.getName()).log(Level.SEVERE, null, ex);
            throw ex;           
        }
    }
    
    public static void main(String [] args) throws Exception {
        String inputPath = "D://JEtest/inputR.xls";
        String outputPath = "D://JEtest/outputW.xls";
        copyReadWrite(inputPath,outputPath);

    }
    
}
