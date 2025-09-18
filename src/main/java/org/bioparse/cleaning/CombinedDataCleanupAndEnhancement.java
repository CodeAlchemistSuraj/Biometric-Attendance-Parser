package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

public class CombinedDataCleanupAndEnhancement {

    // Helper: convert column number to Excel letter
    public static String columnLetter(int colNum) {
        return CellReference.convertNumToColString(colNum - 1);
    }

    public void removeImages(XSSFWorkbook workbook) {
        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            XSSFSheet sheet = workbook.getSheetAt(s);
            if (sheet.getDrawingPatriarch() != null) {
                sheet.getDrawingPatriarch().getShapes().clear();
            }
        }
    }

    public void removeBlankRowsFromWorkDuration(XSSFWorkbook workbook) {
        Sheet workDuration = workbook.getSheet("WorkDuration");
        if (workDuration != null) {
            for (int r = workDuration.getLastRowNum(); r >= 0; r--) {
                Row row = workDuration.getRow(r);
                if (row == null || isRowBlank(row)) {
                    removeRow(workDuration, r);
                }
            }
        }
    }

    public void processMainSheet(Sheet ws) {
        // Remove rows with "Department:" in column A
        for (int r = ws.getLastRowNum(); r >= 0; r--) {
            Row row = ws.getRow(r);
            if (row != null) {
                Cell c = row.getCell(0);
                if (c != null && c.getStringCellValue().startsWith("Department:")) {
                    removeRow(ws, r);
                }
            }
        }

        // Insert empty row before each "Employee:" row
        for (int r = ws.getLastRowNum(); r >= 0; r--) {
            Row row = ws.getRow(r);
            if (row != null) {
                Cell c = row.getCell(0);
                if (c != null && c.getStringCellValue().startsWith("Employee:")) {
                    insertEmptyRow(ws, r);
                }
            }
        }

        // Remove "Late By", "Early By", "OT" rows from each employee block
        int lastRow = ws.getLastRowNum();
        int r = 1;
        while (r <= lastRow) {
            Row row = ws.getRow(r);
            if (row != null && getCellString(row, 0).startsWith("Employee:")) {
                int blockStart = r + 1;
                int[] toDelete = {blockStart + 4, blockStart + 5, blockStart + 6};
                for (int i = toDelete.length - 1; i >= 0; i--) {
                    if (toDelete[i] <= lastRow) {
                        removeRow(ws, toDelete[i]);
                        lastRow--;
                    }
                }
                r = blockStart + 8 - toDelete.length;
            } else {
                r++;
            }
        }
    }

    public void addAttendanceMetrics(Sheet ws) {
        int lastRow = ws.getLastRowNum();
        int lastCol = ws.getRow(0).getLastCellNum();

        int r = 1;
        while (r <= lastRow) {
            Row statusRow = ws.getRow(r);
            if (statusRow != null && "Status".equalsIgnoreCase(getCellString(statusRow, 0))) {
                int inTimeRow = r + 1;
                int outTimeRow = r + 2;
                int durationRow = r + 3;
                int shiftRow = r + 4;
                int blockEndRow = shiftRow;

                // Find last column with data between Status..Shift rows
                int blockLastCol = 1;
                for (int c = 1; c < lastCol; c++) {
                    for (int rr = r; rr <= blockEndRow; rr++) {
                        if (!getCellString(ws.getRow(rr), c).isEmpty()) {
                            blockLastCol = Math.max(blockLastCol, c + 1);
                        }
                    }
                }

                // Insert 5 new rows
                for (int i = 0; i < 5; i++) {
                    insertEmptyRow(ws, blockEndRow + 1 + i);
                }
                lastRow += 5;

                // Label them
                ws.getRow(blockEndRow + 1).createCell(0).setCellValue("FullDay");
                ws.getRow(blockEndRow + 2).createCell(0).setCellValue("HalfDay");
                ws.getRow(blockEndRow + 3).createCell(0).setCellValue("OTDays");
                ws.getRow(blockEndRow + 4).createCell(0).setCellValue("OTHours");
                ws.getRow(blockEndRow + 5).createCell(0).setCellValue("HalfOTDay");

                // Write formulas
                for (int c = 1; c < blockLastCol; c++) {
                    String dayCell = columnLetter(c + 1) + "1";

                    // Full Day
                    setFormula(ws, blockEndRow + 1, c,
                            "IF(AND(" + addr(r, c) + "=\"P\"," +
                                    addr(inTimeRow, c) + "<>\"\"," +
                                    addr(outTimeRow, c) + "<>\"\"," +
                                    "IF(AND(" + addr(inTimeRow, c) + "<>\"\"," + addr(outTimeRow, c) + "<>\"\")," +
                                    addr(outTimeRow, c) + "-" + addr(inTimeRow, c) + ">=TIME(8,30,0),FALSE)," +
                                    "ISERROR(SEARCH(\"St\"," + dayCell + "))," +
                                    "ISERROR(SEARCH(\"S\"," + dayCell + "))),1,0)"
                    );

                    // Half Day
                    setFormula(ws, blockEndRow + 2, c,
                            "IF(AND(" + addr(blockEndRow + 1, c) + "=0," +
                                    addr(inTimeRow, c) + "<>\"\"," +
                                    addr(outTimeRow, c) + "<>\"\"," +
                                    "IF(AND(" + addr(inTimeRow, c) + "<>\"\"," + addr(outTimeRow, c) + "<>\"\")," +
                                    addr(outTimeRow, c) + "-" + addr(inTimeRow, c) + ">=TIME(4,0,0),FALSE)," +
                                    "OR(" + addr(r, c) + "=\"P\"," + addr(r, c) + "=\"½P\")),0.5,0)"
                    );

                    // Weekly Off / Holiday Present (count as 1 or 0.5)
                    setFormula(ws, blockEndRow + 3, c,
                            "IF(OR(" + addr(r, c) + "=\"WO\"," + addr(r, c) + "=\"H\"," + addr(r, c) + "=\"HP\"," + addr(r, c) + "=\"WOP\")," +
                                    "IF(AND(" + addr(inTimeRow, c) + "<>\"\"," + addr(outTimeRow, c) + "<>\"\"),1," +
                                    "IF(OR(" + addr(inTimeRow, c) + "<>\"\"," + addr(outTimeRow, c) + "<>\"\"),0.5,0)),0)"
                    );

                    // Overtime
                    setFormula(ws, blockEndRow + 4, c,
                            "IF(AND(" + addr(r, c) + "=\"P\"," +
                                    addr(inTimeRow, c) + "<>\"\"," +
                                    addr(outTimeRow, c) + "<>\"\")," +
                                    "IF(AND(" + addr(inTimeRow, c) + "<>\"\"," + addr(outTimeRow, c) + "<>\"\")," +
                                    "MAX(0,(" + addr(outTimeRow, c) + "-" + addr(inTimeRow, c) + ")-(8.5/24)),0),0)"
                    );

                    // WO/H/HP/WOP but only one of In or Out filled (count as 1)
                    setFormula(ws, blockEndRow + 5, c,
                            "IF(OR(" + addr(r, c) + "=\"WO\", " + addr(r, c) + "=\"H\", " + addr(r, c) + "=\"HP\", " + addr(r, c) + "=\"WOP\"), " +
                                    "IF(OR(AND(" + addr(inTimeRow, c) + "<>\"\", " + addr(outTimeRow, c) + "=\"\"), " +
                                    "AND(" + addr(inTimeRow, c) + "=\"\", " + addr(outTimeRow, c) + "<>\"\")), 1, 0), 0)"
                    );
                }


                r = blockEndRow + 6;
            } else {
                r++;
            }
        }
    }

    private static String getCellString(Row row, int col) {
        if (row == null) return "";
        Cell cell = row.getCell(col);
        if (cell == null) return "";
        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue().trim();
    }

    private static void removeRow(Sheet sheet, int rowIndex) {
        // Remove any merged regions that include the row to be deleted
        if (sheet instanceof XSSFSheet) {
            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            for (int i = xssfSheet.getNumMergedRegions() - 1; i >= 0; i--) {
                CellRangeAddress region = xssfSheet.getMergedRegion(i);
                if (region.getFirstRow() <= rowIndex && region.getLastRow() >= rowIndex) {
                    xssfSheet.removeMergedRegion(i);
                }
            }
        }

        // Proceed with row removal
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        } else if (rowIndex == lastRowNum) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    private static void insertEmptyRow(Sheet sheet, int rowIndex) {
        // Remove any merged regions that include the insertion point
        if (sheet instanceof XSSFSheet) {
            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            for (int i = xssfSheet.getNumMergedRegions() - 1; i >= 0; i--) {
                CellRangeAddress region = xssfSheet.getMergedRegion(i);
                if (region.getFirstRow() <= rowIndex && region.getLastRow() >= rowIndex) {
                    xssfSheet.removeMergedRegion(i);
                }
            }
        }

        // Insert the empty row
        sheet.shiftRows(rowIndex, sheet.getLastRowNum(), 1);
        sheet.createRow(rowIndex);
    }

    private static boolean isRowBlank(Row row) {
        for (Cell c : row) {
            if (c != null && c.getCellType() != CellType.BLANK && !c.toString().trim().isEmpty())
                return false;
        }
        return true;
    }

    private static void setFormula(Sheet sheet, int rowIdx, int colIdx, String formula) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) row = sheet.createRow(rowIdx);
        Cell cell = row.createCell(colIdx);
        cell.setCellFormula(formula);
    }

    private static String addr(int row, int col) {
        return columnLetter(col + 1) + (row + 1);
    }
}