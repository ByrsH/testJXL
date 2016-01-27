/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package test;

/**
 *
 * @author Administrator
 */

import java.io.*;
import java.util.Date;
import jxl.*;

public class gettypevalue {
    public static void main(String args[]){
        try {
            Workbook book = Workbook.getWorkbook(new File("D://JEtest/测试.xls"));
            Sheet sheet = book.getSheet(0);
            Cell cell1 = sheet.getCell(1, 3);
            System.out.println(cell1.getType());
            String Labelcell1 = "";
            if (cell1.getType() == CellType.STRING_FORMULA) {
                StringFormulaCell StrCell = (StringFormulaCell)cell1;
                String strcell1 = StrCell.getString();
            }
            else if(cell1.getType() == CellType.LABEL) {
                LabelCell LCell = (LabelCell)cell1;
                Labelcell1 = LCell.getString();
            }
            else if(cell1.getType() == CellType.BOOLEAN) {
                BooleanCell LCell = (BooleanCell)cell1;
                Boolean boolcell1 = LCell.getValue();
            }
             else if(cell1.getType() == CellType.BOOLEAN_FORMULA) {
                BooleanFormulaCell LCell = (BooleanFormulaCell)cell1;
                Boolean boolcell1 = LCell.getValue();
            }
             else if(cell1.getType() == CellType.DATE) {
                DateCell LCell = (DateCell)cell1;
                Date datecell1 = LCell.getDate();
            }
            else if(cell1.getType() == CellType.DATE_FORMULA) {
                DateFormulaCell LCell = (DateFormulaCell)cell1;
                Date datecell1 = LCell.getDate();
            }
            else if(cell1.getType() == CellType.EMPTY) {
                String strcell = cell1.getContents();
            }
            else if(cell1.getType() == CellType.ERROR) {
                ErrorCell LCell = (ErrorCell)cell1;
                int intcell1 = LCell.getErrorCode();
            }
            else if(cell1.getType() == CellType.FORMULA_ERROR) {
                ErrorFormulaCell LCell = (ErrorFormulaCell)cell1;
                int intcell1 = LCell.getErrorCode();
            }
            else if(cell1.getType() == CellType.NUMBER) {
                NumberCell LCell = (NumberCell)cell1;
                Double doublecell1 = LCell.getValue();
            }
            else if(cell1.getType() == CellType.NUMBER_FORMULA) {
                NumberFormulaCell LCell = (NumberFormulaCell)cell1;
                Double doublecell1 = LCell.getValue();
            }
            else{
                String result = cell1.getContents();
            }
            
    
            System.out.println("cell1 = " + Labelcell1);
            
        }catch (Exception e) {
            System.out.println(e);
        }
    }
    
}
