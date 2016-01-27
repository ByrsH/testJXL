/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tcell;

/**
 *
 * @author Administrator
 */

import jxl.*;
import java.io.*;

public class TCell {
    public static void main(String [] args) {
        try {
            File file = new File("D://JEtest/测试.xls");
            
            Workbook book = Workbook.getWorkbook(file);
            Sheet sheet = book.getSheet(0);
            Cell cell = sheet.getCell(0, 0);
            
            //返回单元格的行数
            int rows = cell.getRow();
            System.out.println("该单元格所在的行数： " + rows);
            
            //返回单元格的列数
            int col = cell.getColumn();
            System.out.println("该单元格所在的列数： " + col);
            
            //返回单元格内容类型
            CellType cellType = cell.getType();
            System.out.println("该单元格内容类型： " + cellType.toString());
            
            //返回这个单元格是否被隐藏
            boolean judgeHidden = cell.isHidden();
            System.out.println("该单元格是否被隐藏： " + judgeHidden);
            
            //以字符串的形式返回单元格的内容，对于内容更复杂的操作，把它转换成相应的子接口是有必要的。
            String cellContent = cell.getContents();
            System.out.println("该单元格的内容： " + cellContent);
            
            //得到应用于这个单元格的格式，如果该单元格的类型是EMPTY，则返null。
            //一些空的单元格（比如：在模板电子表格上）可能有单元格类型为EMPTY,但是它实际上包含了格式信息。
            jxl.format.CellFormat cellFormat = cell.getCellFormat();
            
            //得到任何特殊单元格属性，例如注释，或者单元格校验当前单元格。如果没有特殊属性，则返回null.
            CellFeatures cellFeatures = cell.getCellFeatures();
            
            
            
            book.close();
            
        }catch (Exception e) {
            System.out.println("Exception: " + e);
        }
    }
    
}
