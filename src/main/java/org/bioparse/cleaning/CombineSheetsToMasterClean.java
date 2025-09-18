package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Scanner;

public class CombineSheetsToMasterClean {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        try {
            System.out.print("Enter Excel file path: ");
            String filePath = scanner.nextLine().trim();

            System.out.print("Enter START column letter (e.g. A): ");
            String startColLetter = scanner.nextLine().trim().toUpperCase();
            int startCol = CellReference.convertColStringToIndex(startColLetter);

            System.out.print("Enter END column letter (e.g. Z): ");
            String endColLetter = scanner.nextLine().trim().toUpperCase();
            int endCol = CellReference.convertColStringToIndex(endColLetter);

            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);

            // Create a master sheet
            Sheet masterSheet = workbook.createSheet("Master");
            int masterRowNum = 0;

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet.getSheetName().equalsIgnoreCase("Master"))
                    continue;

                for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue; // skip completely empty rows

                    boolean hasData = false;
                    for (int c = startCol; c <= endCol; c++) {
                        Cell cell = row.getCell(c);
                        if (cell != null && cell.getCellType() != CellType.BLANK) {
                            hasData = true;
                            break;
                        }
                    }
                    if (!hasData) continue;

                    Row masterRow = masterSheet.createRow(masterRowNum++);
                    int mc = 0;
                    for (int c = startCol; c <= endCol; c++) {
                        Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        Cell newCell = masterRow.createCell(mc++);
                        if (cell != null) {
                            copyCellValue(cell, newCell);
                        }
                    }
                }
            }

            fis.close();

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
            workbook.close();

            System.out.println("Sheets combined into 'Master' successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void copyCellValue(Cell oldCell, Cell newCell) {
        switch (oldCell.getCellType()) {
            case STRING -> newCell.setCellValue(oldCell.getStringCellValue());
            case NUMERIC -> newCell.setCellValue(oldCell.getNumericCellValue());
            case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
            case FORMULA -> newCell.setCellFormula(oldCell.getCellFormula());
            case ERROR -> newCell.setCellErrorValue(oldCell.getErrorCellValue());
            default -> newCell.setBlank();
        }
        if (oldCell.getCellStyle() != null) {
            newCell.setCellStyle(oldCell.getCellStyle());
        }
    }
}
