/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package trange;

/**
 *
 * @author Administrator
 */

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import jxl.Workbook;
import jxl.Cell;
import jxl.Range;
import jxl.read.biff.BiffException;

public class TRange {
    public static void main(String [] args) throws IOException {
        try {
             File file = new File("D://JEtest/测试.xls");
             Workbook book = Workbook.getWorkbook(file);
             
             String [] rangeNameList = book.getRangeNames();
             System.out.println("The range name list:  " + Arrays.toString(rangeNameList));
             
             Range [] rangeList = book.findByName("name");
             System.out.println("The range's name is name number:  " + rangeList.length);
             
             //得到这个范围内的左上方的单元格
             Cell topLeftCell = rangeList[0].getTopLeft();
             System.out.println("左上方单元格的内容，以字符串形式返回： " + topLeftCell.getContents());
             
             //得到这个范围内的右下方的单元格
             Cell bottomRightCell = rangeList[0].getBottomRight();
             System.out.println("右下方单元格的内容，以字符串形式返回： " + bottomRightCell.getContents());
             
             
            //得到这个范围内第一个工作表的索引          ???????????
             int firstSheetIndex = rangeList[0].getFirstSheetIndex();
             System.out.println("这个范围内第一个工作表的索引： " + firstSheetIndex);
             
             //得到这个范围内最后一个工作表的索引       ???????????
             int lastSheetIndex = rangeList[0].getLastSheetIndex();
             System.out.println("这个范围内最后一个工作表的索引： " + lastSheetIndex);
             
        }catch (IOException | BiffException e) {
            System.out.println("Exception:  " + e);
        }
       
        
    }
    
}
