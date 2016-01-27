/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tdemo;

/**
 *
 * @author Administrator
 */

import java.io.IOException;
import jxl.demo.Write;
import jxl.write.WriteException;

public class TWrite {
    public static void main(String [] args) throws IOException, WriteException  {
        
        //写一个电子表格，这个demo 说明了JExcelAPI的大多数特性。例如： text, numbers, fonts, number formats and date formats 
        
        String fn = "D://JEtest/tWrite.xls";
        Write write = new Write(fn);
        //Uses the JExcelAPI to create a spreadsheet 
        write.write();
    }
    
}


//Exception in thread "main" java.io.FileNotFoundException: resources\wealdanddownland.png (系统找不到指定的路径。)
//在源代码中jexcelapi_2_6_12\jexcelapi\resources\wealdanddownland.png有该文件，不知道是因为相对路径写错了，还是jar包里没有该文件。