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

    public static class MergeResult {
        public List<EmployeeData> allEmployees;
        public int monthDays;
        public int year;
        public int month;
        public MergeResult(List<EmployeeData> allEmployees, int monthDays, int year, int month) {
            this.allEmployees = allEmployees;
            this.monthDays = monthDays;
            this.year = year;
            this.month = month;
        }
    }

    public static MergeResult merge(String inputFilePath, String outputFilePath, List<Integer> holidays) throws Exception {
        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook inWb = new XSSFWorkbook(fis);

        int monthDays = 31, year = Calendar.getInstance().get(Calendar.YEAR), month = 1;

        // Improved month/year detection
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
                        month = monthNameToNumber(monthName);
                        monthDays = YearMonth.of(year, month).lengthOfMonth();
                        break outer;
                    }
                }
            }
        }
        System.out.println("Detected: Year=" + year + ", Month=" + month + ", Days=" + monthDays);

        Workbook outWb = new XSSFWorkbook();
        Sheet outSheet = outWb.createSheet("Master");

        CellStyle borderStyle = outWb.createCellStyle();
        borderStyle.setBorderTop(BorderStyle.THIN);
        borderStyle.setBorderBottom(BorderStyle.THIN);
        borderStyle.setBorderLeft(BorderStyle.THIN);
        borderStyle.setBorderRight(BorderStyle.THIN);

        List<EmployeeData> allEmployees = new ArrayList<>();
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

                if (firstVal.isEmpty() || firstVal.startsWith("Department:") || firstVal.startsWith("Monthly Status Report") || firstVal.matches("^[A-Za-z]{3} .*\\d{4}.*")) continue;

                // Start employee block
                if (firstVal.equalsIgnoreCase("Employee:")) {
                    if (current != null) {
                        allEmployees.add(current);
                        outRowNum = writeEmployee(current, outSheet, borderStyle, outRowNum, monthDays, holidays, year, month) + 3;
                    }

                    String empRaw = null;
                    for (int i = 1; i <= 3; i++) {
                        Cell c = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (c != null && !c.toString().trim().isEmpty()) {
                            empRaw = c.toString().trim();
                            break;
                        }
                    }

                    current = new EmployeeData();
                    if (empRaw != null) {
                        String[] parts = empRaw.split(":", 2);
                        current.empId = parts[0].trim();
                        current.empName = (parts.length > 1 ? parts[1].trim() : current.empId);
                    }

                    continue;
                }

                if (current == null) continue;

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
                    allEmployees.add(current);
                    outRowNum = writeEmployee(current, outSheet, borderStyle, outRowNum, monthDays, holidays, year, month) + 3;
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
        return new MergeResult(allEmployees, monthDays, year, month);
    }

    private static int writeEmployee(EmployeeData emp, Sheet outSheet, CellStyle borderStyle,
                                     int startRow, int days, List<Integer> holidays, int year, int month) {

        int rowNum = startRow;
        int maxDataIndex = emp.dailyData.values().stream()
                .mapToInt(list -> {
                    for (int i = list.size() - 1; i >= 0; i--) {
                        if (!list.get(i).isEmpty()) return i;
                    }
                    return -1;
                })
                .max().orElse(-1);

        int colsToWrite = maxDataIndex + 1;

        Row empRow = outSheet.createRow(rowNum++);
        createCell(empRow, 0, "Employee:", borderStyle);
        createCell(empRow, 1, emp.empId + " : " + emp.empName, borderStyle);

        for (Map.Entry<String, List<String>> e : emp.dailyData.entrySet()) {
            Row r = outSheet.createRow(rowNum++);
            createCell(r, 0, e.getKey(), borderStyle);

            int colIndex = 1;
            for (int i = 0; i < colsToWrite; i++) {
                int dayNum = i + 1;
                if (isWeekend(year, month, dayNum) || (holidays != null && holidays.contains(dayNum))) {
                    continue;
                }
                createCell(r, colIndex++, e.getValue().get(i), borderStyle);
            }
        }
        return rowNum - 1; // Return last row index
    }

    private static boolean isWeekend(int year, int month, int day) {
        Calendar cal = Calendar.getInstance();
        cal.set(year, month - 1, day);
        int dow = cal.get(Calendar.DAY_OF_WEEK);
        return dow == Calendar.SATURDAY || dow == Calendar.SUNDAY;
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
}