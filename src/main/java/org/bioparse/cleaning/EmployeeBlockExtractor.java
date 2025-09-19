package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class EmployeeBlockExtractor {

    public static void main(String[] args) {
        String inputPath = "C:/Users/MOHD GULZAR KHAN/Desktop/Suraj Working/Java Git/Biomteric Sheets/WorkDurationReport.xlsx";
        String outputPath = "C:/Users/MOHD GULZAR KHAN/Desktop/Suraj Working/Java Git/outputV2.xlsx";

        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook inputWb = new XSSFWorkbook(fis);
             Workbook outputWb = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputPath)) {

            Sheet outputSheet = outputWb.createSheet("Employee Summary");

            // Style: thin black borders
            CellStyle borderStyle = outputWb.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);

            int outRowIndex = 0;

            for (int i = 0; i < inputWb.getNumberOfSheets(); i++) {
                Sheet sheet = inputWb.getSheetAt(i);
                if (sheet.getSheetName().equalsIgnoreCase("Master")) continue;

                boolean insideEmployee = false;
                List<Row> buffer = new ArrayList<>();

                for (Row row : sheet) {
                    if (row == null) continue;

                    Cell firstCell = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String firstVal = (firstCell != null) ? firstCell.toString().trim() : "";

                    // Ignore unwanted header junk
                    if (firstVal.equalsIgnoreCase("Monthly Status Report (Detailed Work Duration)")
                            || firstVal.toLowerCase().startsWith("company:")
                            || firstVal.matches("(?i)^\\w{3}\\s+\\d{2}.*\\d{4}.*")) {
                        continue;
                    }

                    if (firstVal.toLowerCase().startsWith("employee:")) {
                        // If we were already inside previous employee, flush it
                        if (!buffer.isEmpty()) {
                            outRowIndex = writeEmployeeBlock(outputSheet, buffer, outRowIndex, borderStyle);
                            buffer.clear();
                            outRowIndex += 2; // 2 blank rows between employees
                        }
                        insideEmployee = true;
                    }

                    // If inside employee block, collect rows until next employee or end
                    if (insideEmployee) {
                        boolean hasAnyData = false;
                        for (Cell c : row) {
                            if (c != null && c.getCellType() != CellType.BLANK && !c.toString().trim().isEmpty()) {
                                hasAnyData = true;
                                break;
                            }
                        }
                        if (hasAnyData) buffer.add(row);
                    }
                }

                // Flush last employee if exists
                if (!buffer.isEmpty()) {
                    outRowIndex = writeEmployeeBlock(outputSheet, buffer, outRowIndex, borderStyle);
                    buffer.clear();
                    outRowIndex += 2;
                }
            }

            outputWb.write(fos);
            System.out.println("✅ Output saved at: " + outputPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static int writeEmployeeBlock(Sheet outSheet, List<Row> blockRows,
                                          int outRowIndex, CellStyle style) {

        if (blockRows.isEmpty()) return outRowIndex;

        // --- Employee summary row ---
        Row empRow = blockRows.get(0);
        Row outEmpRow = outSheet.createRow(outRowIndex++);
        for (int c = 0; c < empRow.getLastCellNum(); c++) {
            Cell inCell = empRow.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (inCell == null || inCell.toString().trim().isEmpty()) continue;
            Cell newCell = outEmpRow.createCell(c);
            copyCellValue(inCell, newCell);
            newCell.setCellStyle(style);
        }

        // --- Day-wise rows ---
        for (int i = 1; i < blockRows.size(); i++) {
            Row inRow = blockRows.get(i);
            Row outRow = outSheet.createRow(outRowIndex++);
            int outCol = 0;
            for (int c = 0; c < inRow.getLastCellNum(); c++) {
                Cell inCell = inRow.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (inCell == null || inCell.toString().trim().isEmpty()) continue; // skip blank columns
                Cell newCell = outRow.createCell(outCol++);
                copyCellValue(inCell, newCell);
                newCell.setCellStyle(style);
            }
        }

        return outRowIndex;
    }

    private static void copyCellValue(Cell from, Cell to) {
        if (from == null) return;
        switch (from.getCellType()) {
            case STRING -> to.setCellValue(from.getStringCellValue());
            case NUMERIC -> to.setCellValue(from.getNumericCellValue());
            case BOOLEAN -> to.setCellValue(from.getBooleanCellValue());
            case FORMULA -> to.setCellFormula(from.getCellFormula());
            default -> to.setCellValue(from.toString());
        }
    }
}
