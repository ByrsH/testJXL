/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package example;

/**
 *
 * @author Administrator
 */
import java.io.*; 
import jxl.*; 
import jxl.format.UnderlineStyle;
import jxl.write.*; 
 
public class TWrite { 
  public static void main(String[] args) throws IOException, WriteException { 
    try { 
      // 构建Workbook对象, 只读Workbook对象 
      // Method 1：创建可写入的Excel工作薄 
      jxl.write.WritableWorkbook wwb = Workbook.createWorkbook(new File( 
          "D://JEtest/targetfile.xls")); 
 
      // Method 2：将WritableWorkbook直接写入到输出流 
      /* 
       * OutputStream os = new FileOutputStream(targetfile); 
       * jxl.write.WritableWorkbook wwb = Workbook.createWorkbook(os); 
       */ 
 
      // 创建Excel工作表 
      jxl.write.WritableSheet ws = wwb.createSheet("Test Sheet 1", 0); 
 
      // 现在要做的只是实例化API所提供的Excel基本数据类型，并将它们添加到工作表中就可以了 
 
      // 1.添加Label对象 
      jxl.write.Label labelC = new jxl.write.Label(0, 0, 
          "This is a Label cell"); 
      ws.addCell(labelC); 
 
      // 添加带有字型Formatting的对象 
      jxl.write.WritableFont wf = new jxl.write.WritableFont( 
          WritableFont.TIMES, 18, WritableFont.BOLD, true); 
      jxl.write.WritableCellFormat wcfF = new jxl.write.WritableCellFormat( 
          wf); 
      jxl.write.Label labelCF = new jxl.write.Label(1, 0, 
          "This is a Label Cell", wcfF); 
      ws.addCell(labelCF); 
 
      // 添加带有字体颜色Formatting的对象 
      jxl.write.WritableFont wfc = new jxl.write.WritableFont( 
          WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false, 
          UnderlineStyle.NO_UNDERLINE, jxl.format.Colour.RED); 
      jxl.write.WritableCellFormat wcfFC = new jxl.write.WritableCellFormat( 
          wfc); 
      jxl.write.Label labelCFC = new jxl.write.Label(1, 1, 
          "This is a Label Cell", wcfFC); 
      ws.addCell(labelCFC); 
 
      // 2.添加Number对象 
      jxl.write.Number labelN = new jxl.write.Number(0, 1, 3.1415926); 
      ws.addCell(labelN); 
 
      // 添加带有formatting的Number对象 
      jxl.write.NumberFormat nf = new jxl.write.NumberFormat("#.##"); 
      jxl.write.WritableCellFormat wcfN = new jxl.write.WritableCellFormat( 
          nf); 
      jxl.write.Number labelNF = new jxl.write.Number(1, 1, 3.1415926, 
          wcfN); 
      ws.addCell(labelNF); 
 
      // 3.添加Boolean对象 
      jxl.write.Boolean labelB = new jxl.write.Boolean(0, 2, false); 
      ws.addCell(labelB); 
      // 4.添加DateTime对象 
      jxl.write.DateTime labelDT = new jxl.write.DateTime(0, 3, 
          new java.util.Date()); 
      ws.addCell(labelDT); 
 
      // 添加带有formatting的DateFormat对象 
      jxl.write.DateFormat df = new jxl.write.DateFormat( 
          "dd MM yyyy hh:mm:ss"); 
      jxl.write.WritableCellFormat wcfDF = new jxl.write.WritableCellFormat( 
          df); 
      jxl.write.DateTime labelDTF = new jxl.write.DateTime(1, 3, 
          new java.util.Date(), wcfDF); 
      ws.addCell(labelDTF); 
 
      // 写入Exel工作表 
      wwb.write(); 
 
      // 关闭Excel工作薄对象 
      wwb.close(); 
 
    } catch (IOException | WriteException e) { 
        System.out.println("Exception:  " + e);
            throw e;
    } 
  } 
 
}
