package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.YearMonth;
import java.time.LocalDate;
import java.time.DayOfWeek;
import java.time.format.TextStyle;
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

    public static MergeResult merge(String inputFilePath, String outputFilePath, List<Integer> holidays)
            throws Exception {
        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook inWb;

        if (inputFilePath.toLowerCase().endsWith(".xls")) {
            inWb = new HSSFWorkbook(fis);
        } else {
            inWb = new XSSFWorkbook(fis);
        }

        int monthDays = 31, year = Calendar.getInstance().get(Calendar.YEAR), month = 1;

        // Month/year detection
        outer: for (int s = 0; s < inWb.getNumberOfSheets(); s++) {
            Sheet sheet = inWb.getSheetAt(s);
            for (Row row : sheet) {
                if (row == null)
                    continue;
                for (int col = 0; col < row.getLastCellNum(); col++) {
                    Cell c = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (c != null && c.getCellType() == CellType.STRING) {
                        String val = c.getStringCellValue().trim();

                        boolean matchesMonthDayYear = val.matches("(?i)^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{1,2}\\s+\\d{4}.*");
                        boolean matchesMonthYear = val.matches("(?i)^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{4}.*");

                        if (matchesMonthDayYear || matchesMonthYear) {
                            String cleanVal = val.replaceAll("\\s+", " ").trim();
                            String[] parts = cleanVal.split(" ");
                            String monthName = parts[0];
                            
                            try {
                                int detectedYear = year;
                                
                                if (matchesMonthDayYear && parts.length >= 3) {
                                    detectedYear = Integer.parseInt(parts[2]);
                                } else if (matchesMonthYear && parts.length >= 2) {
                                    detectedYear = Integer.parseInt(parts[1]);
                                }
                                
                                int currentYear = Calendar.getInstance().get(Calendar.YEAR);
                                if (detectedYear >= currentYear - 1 && detectedYear <= currentYear + 1) {
                                    year = detectedYear;
                                    month = monthNameToNumber(monthName);
                                    monthDays = YearMonth.of(year, month).lengthOfMonth();

                                    LocalDate firstDay = LocalDate.of(year, month, 1);
                                    System.out.println("✓ Detected: " + cleanVal + " → Year=" + year + 
                                                     ", Month=" + month + " (" + monthName + "), " +
                                                     "Days=" + monthDays + ", Starts: " + firstDay.getDayOfWeek());
                                    break outer;
                                }
                            } catch (Exception ex) {
                                System.out.println("⚠️  Could not parse date from: " + val);
                            }
                        }
                    }
                }
            }
        }

        Workbook outWb = new XSSFWorkbook();
        Sheet outSheet = outWb.createSheet("Master");

        CellStyle borderStyle = outWb.createCellStyle();
        borderStyle.setBorderTop(BorderStyle.THIN);
        borderStyle.setBorderBottom(BorderStyle.THIN);
        borderStyle.setBorderLeft(BorderStyle.THIN);
        borderStyle.setBorderRight(BorderStyle.THIN);

        CellStyle weekendStyle = outWb.createCellStyle();
        weekendStyle.cloneStyleFrom(borderStyle);
        weekendStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        weekendStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle holidayStyle = outWb.createCellStyle();
        holidayStyle.cloneStyleFrom(borderStyle);
        holidayStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        holidayStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        List<EmployeeData> allEmployees = new ArrayList<>();
        int outRowNum = 0;

        for (int s = 0; s < inWb.getNumberOfSheets(); s++) {
            Sheet sheet = inWb.getSheetAt(s);
            if (sheet.getSheetName().equalsIgnoreCase("Master"))
                continue;

            System.out.println("\n========================================");
            System.out.println("Processing sheet: " + sheet.getSheetName());
            System.out.println("========================================");

            // STEP 1: Unmerge ALL cells in this sheet
            unmergeAllCells(sheet);

            // STEP 2: Remove completely empty columns
            Map<Integer, Integer> columnMapping = removeEmptyColumnsAndGetMapping(sheet);
            
            System.out.println("Column mapping after cleanup: " + columnMapping);

            EmployeeData current = null;

            for (Row row : sheet) {
                if (row == null)
                    continue;
                Cell first = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (first == null || first.getCellType() != CellType.STRING)
                    continue;
                String firstVal = first.getStringCellValue().trim();

                if (firstVal.isEmpty() || firstVal.startsWith("Department:")
                        || firstVal.startsWith("Monthly Status Report") || firstVal.matches("^[A-Za-z]{3} .*\\d{4}.*"))
                    continue;

                // Start employee block
                if (firstVal.equalsIgnoreCase("Employee:")) {
                    if (current != null && hasEssentialData(current)) {
                        allEmployees.add(current);
                        outRowNum = writeEmployee(current, outSheet, borderStyle, weekendStyle, holidayStyle, outRowNum,
                                monthDays, holidays, year, month) + 3;
                    }

                    String empRaw = null;
                    // Search for employee info in next columns (after empty column removal)
                    for (int col = 1; col <= 10; col++) {
                        Cell c = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (c != null) {
                            String cellVal = c.toString().trim();
                            if (!cellVal.isEmpty() && !isSummaryText(cellVal)) {
                                // Look for pattern: number : name or just use first non-empty
                                if (cellVal.matches(".*\\d+.*:.*") || empRaw == null) {
                                    empRaw = cellVal;
                                    if (cellVal.matches(".*\\d+.*:.*")) {
                                        break; // Found the ID:Name pattern
                                    }
                                }
                            }
                        }
                    }

                    current = new EmployeeData();
                    if (empRaw != null) {
                        if (empRaw.contains(":")) {
                            String[] parts = empRaw.split(":", 2);
                            current.empId = parts[0].trim();
                            current.empName = (parts.length > 1 ? parts[1].trim() : current.empId);
                        } else {
                            current.empId = empRaw;
                            current.empName = empRaw;
                        }
                        System.out.println("\n✓ Found employee: " + current.empId + " : " + current.empName);
                    } else {
                        System.out.println("\n⚠️  Could not find employee info in row");
                    }
                    continue;
                }

                if (current == null)
                    continue;

                String label = firstVal;

                // Read data from consecutive columns (empty columns already removed)
                List<String> values = new ArrayList<>();
                
                // Start from column 1 and read exactly monthDays values
                for (int col = 1; col <= monthDays; col++) {
                    Cell c = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String cellValue = "";
                    
                    if (c != null) {
                        cellValue = c.toString().trim();
                        // Skip summary text
                        if (isSummaryText(cellValue)) {
                            cellValue = "";
                        }
                    }
                    
                    values.add(cellValue);
                }

                // Ensure exact monthDays
                while (values.size() < monthDays) {
                    values.add("");
                }
                if (values.size() > monthDays) {
                    values = values.subList(0, monthDays);
                }

                // Only store valid data types
                if (isValidDataType(label)) {
                    current.dailyData.put(label, values);
                    
                    // Count non-empty values for verification
                    long nonEmptyCount = values.stream().filter(v -> !v.isEmpty()).count();
                    System.out.println("  " + label + ": " + nonEmptyCount + " non-empty values out of " + monthDays + " days");
                }

                // Finalize employee when we have Shift data
                if (label.equalsIgnoreCase("Shift")) {
                    if (hasEssentialData(current)) {
                        allEmployees.add(current);
                        outRowNum = writeEmployee(current, outSheet, borderStyle, weekendStyle, holidayStyle, outRowNum,
                                monthDays, holidays, year, month) + 3;
                        current = null;
                    }
                }
            }

            // Handle any remaining employee data
            if (current != null && hasEssentialData(current)) {
                allEmployees.add(current);
                outRowNum = writeEmployee(current, outSheet, borderStyle, weekendStyle, holidayStyle, outRowNum,
                        monthDays, holidays, year, month) + 3;
            }
        }

        inWb.close();
        fis.close();

        FileOutputStream fos = new FileOutputStream(outputFilePath);
        outWb.write(fos);
        fos.close();
        outWb.close();

        System.out.println("\n========================================");
        System.out.println("✓ Master file created at: " + outputFilePath);
        System.out.println("✓ Total employees processed: " + allEmployees.size());
        System.out.println("========================================");
        
        return new MergeResult(allEmployees, monthDays, year, month);
    }

    /**
     * Unmerge all cells in a sheet first
     */
    private static void unmergeAllCells(Sheet sheet) {
        int mergedRegionsCount = sheet.getNumMergedRegions();
        for (int i = mergedRegionsCount - 1; i >= 0; i--) {
            sheet.removeMergedRegion(i);
        }
        System.out.println("Step 1: Unmerged " + mergedRegionsCount + " regions");
    }

    /**
     * CRITICAL METHOD: Remove completely empty columns from the sheet.
     * A column is considered empty if ALL rows in important data rows are empty.
     * Returns a mapping of old column index to new column index.
     */
    private static Map<Integer, Integer> removeEmptyColumnsAndGetMapping(Sheet sheet) {
        // Find max column count
        int maxCol = 0;
        for (Row row : sheet) {
            if (row != null && row.getLastCellNum() > maxCol) {
                maxCol = row.getLastCellNum();
            }
        }

        // Find all employee data rows (rows starting with valid labels)
        List<Row> dataRows = new ArrayList<>();
        for (Row row : sheet) {
            if (row != null) {
                Cell first = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (first != null && first.getCellType() == CellType.STRING) {
                    String label = first.getStringCellValue().trim();
                    if (label.equalsIgnoreCase("Employee:") || isValidDataType(label)) {
                        dataRows.add(row);
                    }
                }
            }
        }

        System.out.println("Step 2: Analyzing " + maxCol + " columns across " + dataRows.size() + " data rows");

        // Identify empty columns (columns that have NO data in ANY data row)
        Set<Integer> emptyColumns = new HashSet<>();
        for (int col = 1; col < maxCol; col++) { // Start from col 1, preserve col 0 (labels)
            boolean hasData = false;
            
            for (Row row : dataRows) {
                Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null) {
                    String cellValue = cell.toString().trim();
                    if (!cellValue.isEmpty() && !isSummaryText(cellValue)) {
                        hasData = true;
                        break;
                    }
                }
            }
            
            if (!hasData) {
                emptyColumns.add(col);
            }
        }

        System.out.println("Found " + emptyColumns.size() + " completely empty columns: " + emptyColumns);

        // If no empty columns, return identity mapping
        if (emptyColumns.isEmpty()) {
            Map<Integer, Integer> identityMap = new HashMap<>();
            for (int i = 0; i < maxCol; i++) {
                identityMap.put(i, i);
            }
            return identityMap;
        }

        // Create new column mapping
        Map<Integer, Integer> columnMapping = new HashMap<>();
        int newColIndex = 0;
        for (int oldCol = 0; oldCol < maxCol; oldCol++) {
            if (!emptyColumns.contains(oldCol)) {
                columnMapping.put(oldCol, newColIndex);
                newColIndex++;
            }
        }

        // Now physically remove the empty columns by shifting cells
        List<Integer> sortedEmptyColumns = new ArrayList<>(emptyColumns);
        Collections.sort(sortedEmptyColumns, Collections.reverseOrder()); // Process from right to left

        for (int emptyCol : sortedEmptyColumns) {
            // Shift all cells to the right of emptyCol to the left
            for (Row row : sheet) {
                if (row != null) {
                    // Shift cells left
                    for (int col = emptyCol; col < maxCol - 1; col++) {
                        Cell sourceCell = row.getCell(col + 1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        Cell targetCell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        
                        if (sourceCell != null) {
                            copyCellValue(sourceCell, targetCell);
                        } else {
                            // Clear target if source is null
                            targetCell.setBlank();
                        }
                    }
                    // Remove the last cell (it's now duplicate)
                    Cell lastCell = row.getCell(maxCol - 1);
                    if (lastCell != null) {
                        row.removeCell(lastCell);
                    }
                }
            }
        }

        System.out.println("Step 2 Complete: Removed " + emptyColumns.size() + " empty columns");
        return columnMapping;
    }

    /**
     * Copy cell value from source to target
     */
    private static void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING:
                target.setCellValue(source.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(source)) {
                    target.setCellValue(source.getDateCellValue());
                } else {
                    target.setCellValue(source.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                target.setCellValue(source.getBooleanCellValue());
                break;
            case FORMULA:
                target.setCellFormula(source.getCellFormula());
                break;
            case BLANK:
                target.setBlank();
                break;
            default:
                target.setCellValue(source.toString());
        }
        
        // Copy style if available
        if (source.getCellStyle() != null) {
            target.setCellStyle(source.getCellStyle());
        }
    }

    /**
     * Check if text is summary text that should be ignored
     */
    private static boolean isSummaryText(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }
        
        String lower = value.toLowerCase();
        return lower.contains("total work duration") ||
               lower.contains("total ot") ||
               lower.contains("present:") ||
               lower.contains("absent:") ||
               lower.contains("weeklyoff:") ||
               lower.contains("holidays:") ||
               lower.contains("leaves taken") ||
               lower.contains("late by hrs") ||
               lower.contains("early by hrs") ||
               lower.contains("shift count") ||
               lower.contains("average working") ||
               lower.contains("hrs.");
    }

    private static boolean isValidDataType(String label) {
        String[] validTypes = { "Status", "InTime", "OutTime", "Duration", "Late By", "Early By", "OT", "Shift" };
        for (String type : validTypes) {
            if (type.equalsIgnoreCase(label)) {
                return true;
            }
        }
        return false;
    }

    private static boolean hasEssentialData(EmployeeData emp) {
        boolean hasRequired = emp.dailyData.containsKey("Status") && 
                             emp.dailyData.containsKey("InTime") &&
                             emp.dailyData.containsKey("OutTime") && 
                             emp.dailyData.containsKey("Duration");
        
        if (!hasRequired) {
            System.out.println("  ⚠️  Missing essential data fields");
        }
        
        return hasRequired;
    }

    private static int writeEmployee(EmployeeData emp, Sheet outSheet, CellStyle borderStyle, CellStyle weekendStyle,
            CellStyle holidayStyle, int startRow, int days, List<Integer> holidays, int year, int month) {

        int rowNum = startRow;

        // Create employee header row
        Row empRow = outSheet.createRow(rowNum++);
        createCell(empRow, 0, "Employee:", borderStyle);
        createCell(empRow, 1, emp.empId + " : " + emp.empName, borderStyle);

        // Create day header row
        Row dayHeaderRow = outSheet.createRow(rowNum++);
        createCell(dayHeaderRow, 0, "Day", borderStyle);

        for (int dayNum = 1; dayNum <= days; dayNum++) {
            LocalDate currentDate = LocalDate.of(year, month, dayNum);
            String dayName = currentDate.getDayOfWeek().getDisplayName(TextStyle.SHORT, Locale.ENGLISH);
            String dayLabel = dayNum + " " + dayName;
            CellStyle headerStyle = borderStyle;

            boolean isWeekend = isWeekend(year, month, dayNum);
            boolean isHoliday = holidays != null && holidays.contains(dayNum) && !isWeekend;

            if (isWeekend) {
                dayLabel += " (WE)";
                headerStyle = weekendStyle;
            } else if (isHoliday) {
                dayLabel += " (H)";
                headerStyle = holidayStyle;
            }

            createStyledCell(dayHeaderRow, dayNum, dayLabel, headerStyle);
        }

        // Write data rows
        String[] dataTypes = { "Status", "InTime", "OutTime", "Duration", "Late By", "Early By", "OT", "Shift" };

        for (String dataType : dataTypes) {
            List<String> data = emp.dailyData.get(dataType);
            if (data == null) {
                data = new ArrayList<>();
                for (int i = 0; i < days; i++) {
                    data.add("");
                }
            }

            Row dataRow = outSheet.createRow(rowNum++);
            createCell(dataRow, 0, dataType, borderStyle);

            for (int dayNum = 1; dayNum <= days; dayNum++) {
                int dataIndex = dayNum - 1;
                String originalValue = dataIndex < data.size() ? data.get(dataIndex) : "";
                if (originalValue == null) originalValue = "";

                CellStyle cellStyle = borderStyle;

                boolean isWeekend = isWeekend(year, month, dayNum);
                boolean isHoliday = holidays != null && holidays.contains(dayNum) && !isWeekend;

                if (isWeekend) {
                    cellStyle = weekendStyle;
                } else if (isHoliday) {
                    cellStyle = holidayStyle;
                }

                createStyledCell(dataRow, dayNum, originalValue, cellStyle);
            }
        }

        return rowNum - 1;
    }

    private static boolean isWeekend(int year, int month, int day) {
        try {
            if (day < 1 || day > YearMonth.of(year, month).lengthOfMonth()) {
                return false;
            }
            LocalDate date = LocalDate.of(year, month, day);
            DayOfWeek dayOfWeek = date.getDayOfWeek();
            return dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY;
        } catch (Exception e) {
            return false;
        }
    }

    private static void createCell(Row r, int col, String val, CellStyle style) {
        Cell c = r.createCell(col);
        c.setCellValue(val);
        c.setCellStyle(style);
    }

    private static void createStyledCell(Row r, int col, String val, CellStyle style) {
        Cell c = r.createCell(col);
        c.setCellValue(val);
        c.setCellStyle(style);
    }

    private static int monthNameToNumber(String name) {
        try {
            String shortName = name.substring(0, 3).toUpperCase();
            switch (shortName) {
            case "JAN": return 1;
            case "FEB": return 2;
            case "MAR": return 3;
            case "APR": return 4;
            case "MAY": return 5;
            case "JUN": return 6;
            case "JUL": return 7;
            case "AUG": return 8;
            case "SEP": return 9;
            case "OCT": return 10;
            case "NOV": return 11;
            case "DEC": return 12;
            default: return 1;
            }
        } catch (Exception e) {
            return 1;
        }
    }
    
}