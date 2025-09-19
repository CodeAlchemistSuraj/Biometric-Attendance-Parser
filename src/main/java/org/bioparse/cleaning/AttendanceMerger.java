package org.bioparse.cleaning;



import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.YearMonth;
import java.util.*;

public class AttendanceMerger {

    static class EmployeeData {
        String empId;
        String empName;
        Map<String, List<String>> dailyData = new LinkedHashMap<>();
    }

    public static void merge(String inputFilePath, String outputFilePath) throws Exception {
        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook inWb = new XSSFWorkbook(fis);

        // Detect month + days from any sheet header
        int monthDays = 31, year = 2025;
        outer:
        for (int s = 0; s < inWb.getNumberOfSheets(); s++) {
            Sheet sheet = inWb.getSheetAt(s);
            for (Row row : sheet) {
                if (row == null) continue;
                Cell c = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (c != null && c.getCellType() == CellType.STRING) {
                    String val = c.getStringCellValue().trim();
                    if (val.matches("^[A-Za-z]{3,9} \\d{2} \\d{4}.*")) {
                        String[] parts = val.split(" ");
                        String monthName = parts[0];
                        year = Integer.parseInt(parts[2]);
                        int month = monthNameToNumber(monthName);
                        monthDays = YearMonth.of(year, month).lengthOfMonth();
                        System.out.println("Detected Month: " + monthName + " (" + monthDays + " days)");
                        break outer;
                    }
                }
            }
        }

        Workbook outWb = new XSSFWorkbook();
        Sheet outSheet = outWb.createSheet("Master");

        // Border style
        CellStyle borderStyle = outWb.createCellStyle();
        borderStyle.setBorderTop(BorderStyle.THIN);
        borderStyle.setBorderBottom(BorderStyle.THIN);
        borderStyle.setBorderLeft(BorderStyle.THIN);
        borderStyle.setBorderRight(BorderStyle.THIN);

        int outRowNum = 0;

        for (int s = 0; s < inWb.getNumberOfSheets(); s++) {
            Sheet sheet = inWb.getSheetAt(s);
            if (sheet.getSheetName().equalsIgnoreCase("Master")) continue;

            EmployeeData current = null;

            for (Row row : sheet) {
                if (row == null) continue;

                Cell first = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (first == null || first.getCellType() != CellType.STRING) continue;
                String firstVal = first.getStringCellValue().trim();

                // Skip unwanted junk
                if (firstVal.isEmpty()
                        || firstVal.startsWith("Department:")
                        || firstVal.startsWith("Monthly Status Report")
                        || firstVal.matches("^[A-Za-z]{3} .*\\d{4}.*")) {
                    continue;
                }

                // Start of employee block
                if (firstVal.equalsIgnoreCase("Employee:")) {
                    if (current != null) {
                        writeEmployee(current, outSheet, borderStyle, outRowNum, monthDays);
                        outRowNum = outSheet.getLastRowNum() + 3; // gap of 2 blank rows
                    }

                    String empRaw = row.getCell(1).getStringCellValue();
                    String[] parts = empRaw.split(":", 2);
                    current = new EmployeeData();
                    current.empId = parts[0].trim();
                    current.empName = (parts.length > 1 ? parts[1].trim() : current.empId);
                    continue;
                }

                if (current == null) continue;

                // Any line like Status / InTime / OutTime / Duration / Shift
                String label = firstVal;
                List<String> values = new ArrayList<>();
                for (int i = 1; i < row.getLastCellNum(); i++) {
                    Cell c = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    values.add(c == null ? "" : c.toString().trim());
                }

                if (values.size() > monthDays) values = values.subList(0, monthDays);
                while (values.size() < monthDays) values.add("");

                current.dailyData.put(label, values);

                if (label.equalsIgnoreCase("Shift")) {
                    writeEmployee(current, outSheet, borderStyle, outRowNum, monthDays);
                    outRowNum = outSheet.getLastRowNum() + 3;
                    current = null;
                }
            }
        }

        inWb.close();
        fis.close();

        FileOutputStream fos = new FileOutputStream(outputFilePath);
        outWb.write(fos);
        fos.close();
        outWb.close();

        System.out.println("Master file created at: " + outputFilePath);
    }

    private static void writeEmployee(EmployeeData emp, Sheet outSheet, CellStyle borderStyle, int startRow, int days) {
        int rowNum = startRow;

        // Employee line
        Row empRow = outSheet.createRow(rowNum++);
        createCell(empRow, 0, "Employee:", borderStyle);
        createCell(empRow, 1, emp.empId + " : " + emp.empName, borderStyle);

        for (Map.Entry<String, List<String>> e : emp.dailyData.entrySet()) {
            Row r = outSheet.createRow(rowNum++);
            createCell(r, 0, e.getKey(), borderStyle);
            for (int i = 0; i < days; i++) {
                createCell(r, i + 1, e.getValue().get(i), borderStyle);
            }
        }
    }

    private static void createCell(Row r, int col, String val, CellStyle style) {
        Cell c = r.createCell(col);
        c.setCellValue(val);
        c.setCellStyle(style);
    }

    private static int monthNameToNumber(String name) {
        String shortName = name.substring(0, 3).toUpperCase(Locale.ENGLISH);
        return java.time.Month.valueOf(shortName).getValue();
    }

    public static void main(String[] args) throws Exception {
        String inputFile = "C:\\Users\\MOHD GULZAR KHAN\\Desktop\\Suraj Working\\Java Git\\Biomteric Sheets\\WorkDurationReport.xlsx";
        String outputFile = "C:\\Users\\MOHD GULZAR KHAN\\Desktop\\Suraj Working\\Java Git\\output.xlsx";
        merge(inputFile, outputFile);
    }
}
