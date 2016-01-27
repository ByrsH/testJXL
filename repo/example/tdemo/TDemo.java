/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tdemo;

import java.io.File;
import java.io.IOException;
import jxl.Workbook;
import jxl.demo.Demo;
import jxl.read.biff.BiffException;

/**
 *
 * @author Administrator
 */
public class TDemo {
    public static void main(String [] args) throws Exception {

        try {
            File file = new File("D://JEtest/CSV.xls");
            Workbook book = Workbook.getWorkbook(file);
            Demo demo = new Demo();
//            String [] str = {"-h"};
//            demo.main(str);
//            System.out.println("111111111111");
            String [] str2 = {"-rw","D://JEtest/CSV.xls","D://JEtest/demo.xls"};
            demo.main(str2);
                
            book.close();     
        }catch (IOException | BiffException e){
            System.out.println("Exception:  " + e);
            throw e;
        }finally {
            
        }
        
    }
    
}
