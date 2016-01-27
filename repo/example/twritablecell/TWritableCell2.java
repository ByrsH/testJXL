/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package twritablecell;

/**
 *
 * @author Administrator
 */

import java.io.File;
import java.io.IOException;
import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.format.Format;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class TWritableCell2 {
    public static void main(String [] args) throws IOException, BiffException, IndexOutOfBoundsException, WriteException  {
        Workbook book;
        WritableWorkbook wb;
        try { 
            File file = new File("D://JEtest/测试.xls");
            
            book = Workbook.getWorkbook(file);
            wb = Workbook.createWorkbook(file, book);
            try {
                WritableSheet sheet = wb.getSheet(2);
            WritableCell cell = sheet.getWritableCell(1, 3);
            WritableCell cell2 = sheet.getWritableCell(0, 0);
            
            System.out.println("cell 的内容：  " + cell.getContents());
            
            
            //设置单元格格式
            CellFormat cf = cell.getCellFormat();       
            cell2.setCellFormat(cf);
            
            Label label = new Label(0, 6, "hello",cf);
            sheet.addCell(label);
            
            
            //进行深拷贝，返回的单元格仍然需要添加到工作表上。不自动添加单元格到表上,客户端程序会改变某些属性,如价值或格式
            //参数为新单元格的列、行
            WritableCell copyCell = cell.copyTo(3, 0);
            sheet.addCell(copyCell);
            
            
            //得到单元格属性
            WritableCell cell3 = sheet.getWritableCell(1, 1);
            WritableCellFeatures cellFeature = cell3.getWritableCellFeatures();
            cellFeature.setComment("yrs");
            System.out.println(cellFeature.getComment());
            
            //设置单元格属性
            //cell2.setCellFeatures(cellFeature);
                
            }finally{
                System.out.println("11111111111");
                wb.write();
                wb.close();
                book.close();
            }
        }catch (IOException | BiffException | IndexOutOfBoundsException | WriteException e) {
            System.out.println("Exception:  " + e);
            throw e;      
        }
        
    }
    
}
