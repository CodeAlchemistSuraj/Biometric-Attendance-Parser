package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.*;

public class AttendanceReportGenerator {

    private String shiftStart = "09:15";
    private int fullDayMinutes = 510;
    private int halfDayMinutes = 240;
    private int otThresholdMinutes = 30;
    private int halfOtMaxMinutes = 240;

    static class Metrics {
        int totalWorkingDays;
        int totalLates;
        int totalLeaves;
        int totalFullDays;
        int halfDays;
        int finalHalfDays;
        int totalOTDays;
        int totalOTHours;
        int totalPunchMissed;
        int totalHalfOTDays;
    }

    public void generate(String inputFile, String outputFile,
                         List<AttendanceMerger.EmployeeData> allEmployees,
                         int monthDays,
                         List<Integer> holidays,
                         int year, int month) throws Exception {

        FileInputStream fis = new FileInputStream(inputFile);
        Workbook wb = new XSSFWorkbook(fis);

        Sheet reportSheet = wb.createSheet("Employee_Report");
        int rowNum = 0;

        Row header = reportSheet.createRow(rowNum++);
        String[] cols = {"Employee Code","Total Working Days","Total Lates","Total Leaves",
                "Total Full Days","Half Days","Final HalfDays",
                "Total OT Days","Total OT Hours",
                "Total Punch Missed","Total Half OT Days"};
        for (int i = 0; i < cols.length; i++) {
            header.createCell(i).setCellValue(cols[i]);
        }

        for (AttendanceMerger.EmployeeData emp : allEmployees) {
            Metrics m = computeMetrics(emp, monthDays, holidays, year, month);

            Row row = reportSheet.createRow(rowNum++);
            int c = 0;
            row.createCell(c++).setCellValue(emp.empId);
            row.createCell(c++).setCellValue(m.totalWorkingDays);
            row.createCell(c++).setCellValue(m.totalLates);
            row.createCell(c++).setCellValue(m.totalLeaves);
            row.createCell(c++).setCellValue(m.totalFullDays);
            row.createCell(c++).setCellValue(m.halfDays);
            row.createCell(c++).setCellValue(m.finalHalfDays);
            row.createCell(c++).setCellValue(m.totalOTDays);
            row.createCell(c++).setCellValue(m.totalOTHours);
            row.createCell(c++).setCellValue(m.totalPunchMissed);
            row.createCell(c++).setCellValue(m.totalHalfOTDays);
        }

        for (int i = 0; i < cols.length; i++) {
            reportSheet.autoSizeColumn(i);
        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(outputFile);
        wb.write(fos);
        fos.close();
        wb.close();

        System.out.println("✅ Employee report generated in: " + outputFile);
    }

    public void exportToCsv(String outputCsvPath, List<AttendanceMerger.EmployeeData> allEmployees,
                            int monthDays, List<Integer> holidays, int year, int month) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputCsvPath))) {
            writer.write("Employee Code,Total Working Days,Total Lates,Total Leaves,Total Full Days,Half Days,Final HalfDays,Total OT Days,Total OT Hours,Total Punch Missed,Total Half OT Days\n");
            for (AttendanceMerger.EmployeeData emp : allEmployees) {
                Metrics m = computeMetrics(emp, monthDays, holidays, year, month);
                writer.write(String.format("%s,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d\n",
                        emp.empId, m.totalWorkingDays, m.totalLates, m.totalLeaves, m.totalFullDays,
                        m.halfDays, m.finalHalfDays, m.totalOTDays, m.totalOTHours, m.totalPunchMissed, m.totalHalfOTDays));
            }
        }
        System.out.println("✅ CSV exported to: " + outputCsvPath);
    }

    private Metrics computeMetrics(AttendanceMerger.EmployeeData emp, int monthDays,
                                   List<Integer> holidays, int year, int month) {

        Metrics m = new Metrics();

        List<String> inTimes = emp.dailyData.getOrDefault("InTime", new ArrayList<>());
        List<String> outTimes = emp.dailyData.getOrDefault("OutTime", new ArrayList<>());
        List<String> duration = emp.dailyData.getOrDefault("Duration", new ArrayList<>());
        List<String> status = emp.dailyData.getOrDefault("Status", new ArrayList<>());

        for (int day = 0; day < monthDays; day++) {
            int dayNum = day + 1;
            if (isWeekend(year, month, dayNum) || (holidays != null && holidays.contains(dayNum))) {
                continue;
            }

            String in = AttendanceUtils.safeGet(inTimes, day);
            String out = AttendanceUtils.safeGet(outTimes, day);
            String dur = AttendanceUtils.safeGet(duration, day);
            String stat = AttendanceUtils.safeGet(status, day);

            int minutes = parseMinutes(dur);

            if (!dur.isEmpty() || (!stat.isEmpty())) {
                m.totalWorkingDays++;
            }

            if ("L".equalsIgnoreCase(stat)) {
                m.totalLeaves++;
            }

            if (minutes >= fullDayMinutes) {
                m.totalFullDays++;
            } else if (minutes >= halfDayMinutes) {
                m.halfDays++;
            }

            if (isLate(in)) {
                m.totalLates++;
            }

            if ((in.isEmpty() || out.isEmpty()) && minutes > 0) {
                m.totalPunchMissed++;
            }

            if (minutes > fullDayMinutes + otThresholdMinutes) {
                int otMinutes = minutes - fullDayMinutes;
                m.totalOTDays++;
                m.totalOTHours += otMinutes / 60;
                if (otMinutes < halfOtMaxMinutes) {
                    m.totalHalfOTDays++;
                }
            }
        }

        m.finalHalfDays = m.halfDays / 2;
        return m;
    }

    private static int parseMinutes(String val) {
        if (val == null || val.isEmpty()) return 0;
        try {
            String[] parts = val.split(":");
            int h = Integer.parseInt(parts[0]);
            int m = Integer.parseInt(parts[1]);
            return h * 60 + m;
        } catch (Exception e) {
            return 0;
        }
    }

    private boolean isLate(String inTime) {
        if (inTime == null || inTime.isEmpty()) return false;
        try {
            return java.time.LocalTime.parse(inTime).isAfter(java.time.LocalTime.parse(shiftStart));
        } catch (Exception e) {
            return false;
        }
    }

    private static boolean isWeekend(int year, int month, int day) {
        try {
            // Validate date before creating it
            if (day < 1 || day > java.time.YearMonth.of(year, month).lengthOfMonth()) {
                return false;
            }
            LocalDate d = LocalDate.of(year, month, day);
            DayOfWeek dow = d.getDayOfWeek();
            return dow == DayOfWeek.SATURDAY || dow == DayOfWeek.SUNDAY;
        } catch (Exception e) {
            return false;
        }
    }

    public String getShiftStart() { return shiftStart; }
    public void setShiftStart(String shiftStart) { this.shiftStart = shiftStart; }
}