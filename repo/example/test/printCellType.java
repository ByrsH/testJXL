/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package test;

import jxl.*;

/**
 *
 * @author Administrator
 */
public class printCellType {
    public static void main(String args[]) {
        System.out.println("BOOLEAN = " + CellType.BOOLEAN.toString());
        System.out.println("BOOLEAN_FORMULA = " + CellType.BOOLEAN_FORMULA.toString());
        System.out.println("DATE = " + CellType.DATE.toString());
        System.out.println("DATE_FORMULA = " + CellType.DATE_FORMULA.toString());
        System.out.println("EMPTY = " + CellType.EMPTY.toString());
        System.out.println("ERROR = " + CellType.ERROR.toString());
        System.out.println("FORMULA_ERROR = " + CellType.FORMULA_ERROR.toString());
        System.out.println("LABEL = " + CellType.LABEL.toString());
        System.out.println("NUMBER = " + CellType.NUMBER.toString());
        System.out.println("NUMBER_FORMULA = " + CellType.NUMBER_FORMULA.toString());
        System.out.println("STRING_FORMULA = " + CellType.STRING_FORMULA.toString());
    }
    
}
