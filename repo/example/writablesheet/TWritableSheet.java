/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package writablesheet;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import jxl.CellView;
import jxl.Range;
import jxl.Workbook;
import jxl.format.PageOrientation;
import jxl.format.PaperSize;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableHyperlink;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 *
 * @author Administrator
 */
public class TWritableSheet {
     public static void main(String [] args) throws IOException, BiffException, IndexOutOfBoundsException, WriteException  {
        Workbook book;
        WritableWorkbook wb;
        File file = new File("D://JEtest/测试.xls");
            
        book = Workbook.getWorkbook(file);
        wb = Workbook.createWorkbook(file, book);
        try {             
            WritableSheet wSheet = wb.getSheet(3);
            
            //添加单元格到工作表中
            Label label = new Label(3, 5, "yrs");
            wSheet.addCell(label);
            
            
            //设置工作表名称
            wSheet.setName("yrs");
            
            
            //设置隐藏。弃用，使用SheetSettings里的。
            //void setHidden(boolean hidden)
            
            //设置是否保护，弃用，使用SheetSettings里的。
            //void setProtected(boolean prot)
            
            
            //设置工作表指定列的宽度。这导致Excel调整整个列。如果已经指定的列视图相关联的信息,然后取而代之的是新的数据
            //第一个参数是列，第二个参数是宽度
            wSheet.setColumnView(0, 15);
                        
            //void setColumnView(int col,int width,CellFormat format)
            //弃用，取而代之的是 CellView    
                        
            //void setColumnView(int col,CellView view)
            //设置这个列的视图
            CellView view = new CellView();
            view.setAutosize(true);
            wSheet.setColumnView(1, view);
                   
            
            //设置指定行的高度
            //void setRowView(int row,int height)throws jxl.write.biff.RowsExceededException
            wSheet.setRowView(0, 500);
            
            //void setRowView(int row,boolean collapsed)throws jxl.write.biff.RowsExceededException
            //设定指定行是否崩溃,true 是显示，false不显示该行
            wSheet.setRowView(10,false);
            
            //void setRowView(int row,int height,boolean collapsed)throws jxl.write.biff.RowsExceededException
            //设置指定行的高度和崩溃属性。false显示该行，true不显示该行
            wSheet.setRowView(5, 400, false);
            
            //void setRowView(int row,CellView view)
            //设置行视图
            CellView view2 = new CellView();
            wSheet.setRowView(19, view2);
            System.out.println(view2.isAutosize() + "    " + view2.isHidden());
            //?????????奇怪，运行之后该行不显示了。？？？？？？？？？？
            
            
            //WritableCell getWritableCell(int column, int row)
            //返回可写单元格
            WritableCell wCell = wSheet.getWritableCell(2, 0);
            System.out.println("单元格的内容：  " +  wCell.getContents());
            
            
            //WritableCell getWritableCell(java.lang.String loc)
            //根据 loc 返回单元格。 例如"A4",考虑到性能问题，尽量少用
            WritableCell wCell2 = wSheet.getWritableCell("A2");
            System.out.println("单元格的内容：  " +  wCell2.getContents());
            
            
            //WritableHyperlink[] getWritableHyperlinks()
            //得到这个表中可写的超链接。返回的超链接可能由用户应用程序修改
            WritableHyperlink [] wHlArr = wSheet.getWritableHyperlinks();
            System.out.println("超链接个数：  " +  wHlArr.length);
            
            
            //void insertRow(int row)
            //插入一个空白行，如果row超出工作表的范围，则不做任何动作
            //row是行索引。
            //wSheet.insertRow(1);
            
            
            //void insertColumn(int col)
            //插入一列。col是列索引.如果col超出工作表的范围，则不做任何动作
            //wSheet.insertColumn(1);
            
            
            //void removeColumn(int col)
            //移除一列，如果col超出工作表的范围，则不做任何动作
            //wSheet.removeColumn(1);
            
            
            //void removeRow(int row) 
            //移除一行，如果row超出工作表的范围，则不做任何动作
            //wSheet.removeRow(1);
            
            
            //Range mergeCells(int col1,int row1,int col2,int row2) throws WriteException,jxl.write.biff.RowsExceededException
            //合并单元格，(col1,row1)是合并单元格左上角的单元格，(col2,row2)是右下角单元格
            Range range = wSheet.mergeCells(0, 0, 2, 2);
            
            
            //void setRowGroup(int row1,int row2, boolean collapsed)throws WriteException,jxl.write.biff.RowsExceededException
            //设置一个行组
            //wSheet.setRowGroup(5, 7, true);
            
            
            //void unsetRowGroup(int row1,int row2)throws WriteException,jxl.write.biff.RowsExceededException
            //????????????????????????????????
            //wSheet.unsetRowGroup(5, 7);
            
            
            //void setColumnGroup(int row1,int row2, boolean collapsed)throws WriteException,jxl.write.biff.RowsExceededException
            //wSheet.setColumnGroup(4, 5, true);
            
            
            //void unsetColumnGroup(int row1,int row2)throws WriteException,jxl.write.biff.RowsExceededException
            //wSheet.unsetColumnGroup(4,5);
            
            
            //拆分合并的单元格
            wSheet.unmergeCells(range);
            
            
            //添加超链接
            URL url = new URL("http://www.baidu.com");
            WritableHyperlink wHyperlink = new WritableHyperlink(8,0,url);
            wSheet.addHyperlink(wHyperlink);
            WritableHyperlink wHyperlink2 = new WritableHyperlink(8,1,url);
            wSheet.addHyperlink(wHyperlink2);
            WritableHyperlink wHyperlink3 = new WritableHyperlink(8,2,url);
            wSheet.addHyperlink(wHyperlink3);
            
            
            //移除超链接
            wSheet.removeHyperlink(wHyperlink2); 
            
            //移除超链接，true 是保留内容，false是不保留
            wSheet.removeHyperlink(wHyperlink3, true);
            
            
            //设置页面设置的细节,  参数是页面方向。   
            //PageOrientation.LANDSCAPE 横向 ,   PageOrientation.PORTRAIT   纵向
            //void setPageSetup(PageOrientation p)
            //wSheet.setPageSetup(PageOrientation.PORTRAIT);
            wSheet.setPageSetup(PageOrientation.LANDSCAPE);
            
            //void setPageSetup(PageOrientation p,double hm,double fm)
            //p - the page orientation
            //hm - the header margin, in inches
            //fm - the footer margin, in inches
            wSheet.setPageSetup(PageOrientation.PORTRAIT, 10, 10);
            
            //void setPageSetup(PageOrientation p,PaperSize ps,double hm,double fm)
            //p - the page orientation
            //ps - the paper size
            //hm - the header margin, in inches
            //fm - the footer margin, in inches
            wSheet.setPageSetup(PageOrientation.PORTRAIT, PaperSize.A4 , 20, 20); 
            
            
            //##########好像没变化？？？？？？
            //void addRowPageBreak(int row)
            //Forces a page break at the specified row 
            wSheet.addRowPageBreak(3); 
            
            //void addColumnPageBreak(int col)
            //Forces a page break at the specified column 
            wSheet.addColumnPageBreak(3); 
            
            
            //void addImage(WritableImage image)
            //添加图片，只支持png类型。
            File image = new File("D://JEtest/abc.png");
            WritableImage wImage = new WritableImage(8,4,5,5,image);
            wSheet.addImage(wImage);
            WritableImage wImage2 = new WritableImage(8,8,5,5,image);
            wSheet.addImage(wImage2);
            
            
            //int getNumberOfImages()
            //得到图片的数量
            int imageNumbers = wSheet.getNumberOfImages();
            System.out.println("工作表中图片的数量：   " + imageNumbers);
            
            //WritableImage getImage(int i)
            //返回指定的图片，索引从0开始
            WritableImage wImage3 = wSheet.getImage(1);
            
            
            //移除图片
            wSheet.removeImage(wImage3); 
            
            
            //void applySharedDataValidation(WritableCell cell,int col,int row) throws WriteException
            //Extend the data validation contained in the specified cell across and downwards. 
            //NOTE: The source cell (top left) must have been added to the sheet prior to this method being called 
            wSheet.applySharedDataValidation(wCell2, 5, 5); 
            
            
            wSheet.removeSharedDataValidation(wCell2); 
            
            
            wb.write();
            wb.close();
            book.close();    

        }catch (IOException | IndexOutOfBoundsException | WriteException e) {
            System.out.println("Exception:  " + e);
            throw e;      
        }
        finally {
//            wb.write();
//            wb.close();
//            book.close();
        }
       
    }
}
