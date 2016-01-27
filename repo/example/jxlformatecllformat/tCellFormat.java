/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package jxlformatecllformat;

/**
 *
 * @author Administrator
 */

import java.io.File;
import java.io.IOException;
import jxl.Workbook;
import jxl.Sheet;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.format.Format;
import jxl.format.Orientation;
import jxl.format.Pattern;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;

public class tCellFormat {
    public static void main(String [] args) {
        try {
            File file = new File("D://JEtest/测试.xls");
            
            Workbook book = Workbook.getWorkbook(file);
            Sheet sheet = book.getSheet(1);
            System.out.println("111");
            
            for(int i = 0;i<=8;i++) {
                //列的索引从0开始
                CellFormat colCellFormat = sheet.getColumnFormat(i);    
                
                //得到单元格格式
                Format format = colCellFormat.getFormat();
               
                //访问应用于单元格的格式的Excel格式字符串，这是excel使用的字符串，与java不等效。
                //返回单元格格式字符串
                String formatStr = format.getFormatString();
                System.out.println("单元格格式字符串： " + formatStr);
            }
            
            CellFormat format2 = sheet.getColumnFormat(0);
            //得到这种格式的字体信息
            Font font = format2.getFont();
            System.out.println("该字体名字： " + font.getName());
            
            
            //判断这个单元的内容是否被包装
            boolean judgeWrap = format2.getWrap();
            System.out.println("这个单元内容是否被包装： " + judgeWrap);
            
            
            //得到单元格水平方向上的校准（居中，左对齐等等)
            Alignment alignment = format2.getAlignment();
            System.out.println("对这个水平校准的描述： " + alignment.getDescription());
            
            //得到垂直方向上的校准
            VerticalAlignment vAlignment = format2.getVerticalAlignment();
            System.out.println("对这个垂直校准的描述： " + vAlignment.getDescription());
            
            
            //得到单元格数据的方向
            Orientation orientation = format2.getOrientation();
            System.out.println("对这个数据定位的描述： " + orientation.getDescription());
            
            
            //BorderLineStyle getBorder(Border border)
            //得到单元格边界线的风格，如果指定边界风格是ALL或NONE,则线风格是NONE。
            BorderLineStyle borderLineStyle = format2.getBorder(Border.BOTTOM);
            System.out.println("单元格边界线的风格的描述： " + borderLineStyle.getDescription());
            
            //BorderLineStyle getBorderLine(Border border)
            BorderLineStyle borderLineStyle2 = format2.getBorder(Border.ALL);
            System.out.println("单元格边界线的风格的描述： " + borderLineStyle2.getDescription());
            
            //Colour getBorderColour(Border border)
            Colour colour = format2.getBorderColour(Border.BOTTOM);
            System.out.println("单元格边界线的颜色描述： " + colour.getDescription());
            
            
            //判断这个单元格格式是否有任何的边界，当合并的一组单元格时，可以用于设置新的边界。
            //TRUE表示有。
            boolean judgeBorder = format2.hasBorders();
            System.out.println("判断这个单元格格式是否有任何的边界： " + judgeBorder);
            
            
            //得到单元格的背景色
            Colour backColour = format2.getBackgroundColour();
            System.out.println("单元格的背景色： " + backColour.getDescription());
            
            
            //单元格格式的模式
            Pattern pattern = format2.getPattern();
            System.out.println("背景模式描述： " + pattern.getDescription());
            
            
            //单元格的缩进文本
            int indentation = format2.getIndentation();
            System.out.println("单元格缩进： " + indentation);
            
            
            //是否缩小以适应标记，TRUE是。
            boolean shrink = format2.isShrinkToFit();
            System.out.println("是否缩小以适应标记： " + shrink);
            
            
            //判断特殊的单元格是否被锁定
            boolean locked = format2.isLocked();
            System.out.println("是否锁定特殊的单元格： " + locked);

            
            book.close();
            
        }catch (IOException | BiffException | IndexOutOfBoundsException e) {
            System.out.println("Exception: " + e);
        }
    }
    
}
