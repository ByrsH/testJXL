/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package twritableworkbook;

/**
 *
 * @author Administrator
 */

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import jxl.Range;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class TWritableWorkbook {
    public static void main(String [] args) throws IOException, BiffException, IndexOutOfBoundsException, WriteException  {
        try {
            Workbook book;
            WritableWorkbook wb;
            
            File file = new File("D://JEtest/测试.xls");
            
            if(file.exists()) {
                //如果文件存在; 
                
                book = Workbook.getWorkbook(file);
                wb = Workbook.createWorkbook(file, book);
                
                
                //得到工作簿中的所有表，以数组形式返回。使用这种方法对于大型工作表可能会导致性能问题。
                WritableSheet wSheetarray[] = wb.getSheets();
                System.out.println("工作簿的工作表数：  " + wSheetarray.length);
                
                //返回工作表数
                int wSheetNumbers = wb.getNumberOfSheets();
                System.out.println("工作簿的工作表数：  " + wSheetNumbers);
                
                
                //返回工作表名，字符串数组
                String [] sheetNamesArray = wb.getSheetNames();
                System.out.println("工作表名：  " + Arrays.toString(sheetNamesArray));
                
                
                //根据索引返回表，以0开始。
                WritableSheet wSheet = wb.getSheet(0);
                System.out.println("第0个工作表名为：  " + wSheet.getName());
                //根据表名返回工作表
                WritableSheet wSheet2 = wb.getSheet("Sheet1");
                
                
                //返回指定的单元格，例如Sheet1!A4.    //注意表名不要写错。    
                //使用该方法会影响性能，谨慎使用。
                WritableCell wCell = wb.getWritableCell("Sheet1!A4");
                System.out.println("1111111111111111");
                System.out.println("sheet1!A4单元格内容为：  " + wCell.getContents());
                
                
                //创建可写工作表
                //WritableSheet wCreatSheet = wb.createSheet("name2", 4);
                //System.out.println("工作簿的工作表数：  " + wb.getNumberOfSheets());
                
                
                //##############有问题################################
                //从不同的工作簿中复制一个工作表
                //WritableSheet importWSheet = wb.importSheet("yrs2", 5, sheetWrite);
                //System.out.println("导入工作表名字  ：  " + importWSheet.getName());
                               
                //复制同一个工作簿中的表
                //wb.copySheet(1, "Sheet2", 3);
                //wb.copySheet("Sheet1", "Sheet2", 2);
                //####################################################
                
                //移除指定的工作表
                //wb.removeSheet(2);
                
                //移动这个工作簿中的表，第一个参数为要移动表的索引，第二个参数是移动的目标位置索引。
                //WritableSheet wSheet3 = wb.moveSheet(1, 2);
                
                
                //指示这个工作簿是否被保护 
                //wb.setProtected(false);
                
                
                //public abstract void setColourRGB(Colour c,int r,int g,int b)
                //设置工作簿的RGB值（红，绿，蓝）  范围在0-255
                wb.setColourRGB(Colour.RED, 255, 0, 255);
                //wb.setColourRGB(Colour.LIGHT_BLUE, 0x76, 0xEE, 0x00);
                jxl.write.WritableFont wfc = new jxl.write.WritableFont( 
                    WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE,jxl.format.Colour.BLUE); 
                WritableCellFormat wcf2 = new WritableCellFormat(wfc);// 单元格样式
                wcf2.setBackground(Colour.RED);
                WritableSheet wSheet3 = wb.getSheet("Sheet2");
                wSheet3.addCell(new Label(1, 3, "测试颜色---自定义#76EE00", wcf2)); 
                                  
                  
                
                //根据 cell或range 的名字返回单元格，如果是range ,则返回左上角的单元格
                WritableCell wCell2 = wb.findCellByName("name");
                System.out.println("单元格内容：  "+ wCell2.getContents());
                
                
                //
                Range [] rangeArr = wb.findByName("search");
                
                //
                String [] rangeNames = wb.getRangeNames();
                
                //清除指定名称的range. 注意移除的这个名字，可能导致使用这个名字的公式计算不正确。
                //wb.removeRangeName("sheet1");
                
                
                //向一个表中添加range.   参数add 是range名称，wSheet3是工作表
                //2,2,4,4 表示第二列到第四列，第二行到第四行的范围。
                wb.addNameArea("add", wSheet3, 2, 2, 4, 4);
                
                
                
                //设置一个新的输出文件，这允许同一工作簿被写入不同输出文件,而无需再次读到任何模板
                //???????????????????源文件里没有内容了。
                //File outputFile = new File("D://JEtest/output.xls");
                //wb.setOutputFile(outputFile); 
                
                book.close();
                wb.write();
                wb.close();
                
            }
            else {
                System.out.println("文件不存在  ");
            }
    
        }catch (IOException | BiffException | IndexOutOfBoundsException | WriteException e) {
            System.out.println("Exception:  " + e);
            throw e;
        }
    }
    
}
