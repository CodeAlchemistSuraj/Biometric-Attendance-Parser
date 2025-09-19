package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Scanner;

public class CombineSheetsToMasterClean {

//    public static void copyCellValue(Cell oldCell, Cell newCell) {
//        switch (oldCell.getCellType()) {
//            case STRING -> newCell.setCellValue(oldCell.getStringCellValue());
//            case NUMERIC -> newCell.setCellValue(oldCell.getNumericCellValue());
//            case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
//            case FORMULA -> newCell.setCellFormula(oldCell.getCellFormula());
//            case ERROR -> newCell.setCellErrorValue(oldCell.getErrorCellValue());
//            default -> newCell.setBlank();
//        }
//        if (oldCell.getCellStyle() != null) {
//            newCell.setCellStyle(oldCell.getCellStyle());
//        }
//    }

    public static void copyCellValue(Cell oldCell, Cell newCell) {
        if (oldCell == null) return;

        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            default:
                break;
        }
    }
}
