package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bioparse.cleaning.AttendanceMerger.EmployeeData;

import java.io.*;
import java.time.*;
import java.util.*;

public class AttendanceReportGenerator {

    // --- Configuration ---
    private String shiftStart = "09:30";
    private String onTimeEnd = "09:46";
    private String halfDayThreshold = "10:15";
    private int fullDayMinutes = 510; // 8.5 hrs
    private int halfDayMinutes = 240; // 4 hrs
    private int otThresholdMinutes = 30;
    private int fullOtMinutes = 480; // 8 hrs
    private int allowedLatesPerMonth = 5;

    public static class Metrics {
    	 public int totalWorkingDays;
    	 public int totalLates;
    	 public int totalAbsent;
    	  public int totalFullDays;
    	  public  int halfDays;

    	  public int halfDaysDueToLate;
    	  public int halfDaysDueToPunchMiss;
    	  public int halfDaysDueToDuration;

    	  public int totalPunchMissed;
        

        // OT tracking
    	  public int workingDayOTMinutes;
    	  public int weekendFullOTMinutes;
    	  public int weekendHalfOTMinutes;
    	  public int weekendHolidayPresentDays;
    	  public int totalOTDays;
    	  public int totalHalfOTDays;
    	  public int totalOTMinutes;
    	  public List<String> remarks = new ArrayList<>();
    }

    // ======================= REPORT GENERATION =======================

    public void generate(String inputFile, String outputFile,
                         List<AttendanceMerger.EmployeeData> allEmployees,
                         int monthDays, List<Integer> holidays,
                         int year, int month) throws Exception {

        Workbook wb = new XSSFWorkbook(new FileInputStream(inputFile));
        Sheet sheet = wb.createSheet("Employee_Report");

        String[] cols = {
                "Employee Code",
                "Total Working Days",
                "Total Full Days",
                "Half Days Due To Duration",
                "Half Days Due To Punch Miss",
                "Half Days Due To Lates",
                "Total Half Days",
                "Total Lates",
                "Total Absent",
                "Total Punch Missed",
                "Weekend/Holiday Present Days",
                "Working Day OT Hours",
                "Weekend/Holiday Full OT Hours",
                "Weekend/Holiday Half OT Hours",
                "Total OT Hours",
                "Remarks"
        };

        Row header = sheet.createRow(0);
        for (int i = 0; i < cols.length; i++) {
            header.createCell(i).setCellValue(cols[i]);
        }

        int rowNum = 1;
        for (AttendanceMerger.EmployeeData emp : allEmployees) {
            Metrics m = computeMetrics(emp, monthDays, holidays, year, month);

            Row r = sheet.createRow(rowNum++);
            int c = 0;

            r.createCell(c++).setCellValue(emp.empId);
            r.createCell(c++).setCellValue(m.totalWorkingDays);
            r.createCell(c++).setCellValue(m.totalFullDays);
            r.createCell(c++).setCellValue(m.halfDaysDueToDuration);
            r.createCell(c++).setCellValue(m.halfDaysDueToPunchMiss);
            r.createCell(c++).setCellValue(m.halfDaysDueToLate);
            r.createCell(c++).setCellValue(m.halfDays);
            r.createCell(c++).setCellValue(m.totalLates);
            r.createCell(c++).setCellValue(m.totalAbsent);
            r.createCell(c++).setCellValue(m.totalPunchMissed);
            r.createCell(c++).setCellValue(m.weekendHolidayPresentDays);

            // Working-day OT (rounded)
            int blocks = (m.workingDayOTMinutes + 29) / 30;
            double workingOTHours = blocks * 0.5;

            double weekendFullOTHours = m.weekendFullOTMinutes / 60.0;
            double weekendHalfOTHours = m.weekendHalfOTMinutes / 60.0;

            double totalOTHours = workingOTHours + weekendFullOTHours + weekendHalfOTHours;

            r.createCell(c++).setCellValue(round2(workingOTHours));
            r.createCell(c++).setCellValue(round2(weekendFullOTHours));
            r.createCell(c++).setCellValue(round2(weekendHalfOTHours));
            r.createCell(c++).setCellValue(round2(totalOTHours));
            r.createCell(c++).setCellValue(String.join("; ", m.remarks));
        }

        for (int i = 0; i < cols.length; i++) sheet.autoSizeColumn(i);

        FileOutputStream fos = new FileOutputStream(outputFile);
        wb.write(fos);
        fos.close();
        wb.close();
    }

    // ======================= CORE LOGIC =======================

    public Metrics computeMetrics(AttendanceMerger.EmployeeData emp, int monthDays,
                                   List<Integer> holidays, int year, int month) {

        Metrics m = new Metrics();

        int weekendDays = 0;
        for (int d = 1; d <= monthDays; d++) if (isWeekend(year, month, d)) weekendDays++;

        int holidayWD = 0;
        if (holidays != null)
            for (Integer h : holidays)
                if (!isWeekend(year, month, h)) holidayWD++;

        m.totalWorkingDays = monthDays - weekendDays - holidayWD;

        List<String> inTimes = emp.dailyData.getOrDefault("InTime", new ArrayList<>());
        List<String> outTimes = emp.dailyData.getOrDefault("OutTime", new ArrayList<>());
        List<String> statusList = emp.dailyData.getOrDefault("Status", new ArrayList<>());

        String[] status = new String[monthDays];
        for (int i = 0; i < monthDays; i++) {
            status[i] = Optional.ofNullable(AttendanceUtils.safeGet(statusList, i))
                    .orElse("").toUpperCase();
        }

        // Sandwich rule
        List<Integer> abs = new ArrayList<>();
        for (int i = 0; i < monthDays; i++) if ("A".equals(status[i])) abs.add(i);
        for (int i = 0; i < abs.size() - 1; i++) {
            for (int j = abs.get(i) + 1; j < abs.get(i + 1); j++) {
                if ("H".equals(status[j]) || "WO".equals(status[j])) status[j] = "A";
            }
        }

        int latesSeen = 0;
        LocalTime shiftStartTime = LocalTime.parse(shiftStart);

        for (int d = 0; d < monthDays; d++) {
            int dayNum = d + 1;
            boolean isHolidayOrWeekend = isWeekend(year, month, dayNum)
                    || (holidays != null && holidays.contains(dayNum));

            String in = AttendanceUtils.safeGet(inTimes, d);
            String out = AttendanceUtils.safeGet(outTimes, d);
            boolean inP = in != null && !in.isEmpty();
            boolean outP = out != null && !out.isEmpty();
            boolean punchMiss = inP ^ outP;

            int minutes = calculateDuration(in, out);

            if ("A".equals(status[d]) && !inP && !outP) m.totalAbsent++;

            if (punchMiss) m.totalPunchMissed++;

            // ---------------- WORKING DAY ----------------
            if (!isHolidayOrWeekend && !"A".equals(status[d])) {

                if (punchMiss) {
                    m.halfDays++;
                    m.halfDaysDueToPunchMiss++;
                    m.remarks.add(dayNum + "th: Half day due to punch miss");
                    continue;
                }

                if (isAfterHalfDayThreshold(in)) {
                    m.halfDays++;
                    m.halfDaysDueToLate++;
                    m.remarks.add(dayNum + "th: Half day due to late arrival (>10:15)");
                    continue;
                }

                if (isLate(in)) {
                    latesSeen++;
                    m.totalLates++;
                    if (latesSeen > allowedLatesPerMonth) {
                        m.halfDays++;
                        m.halfDaysDueToLate++;
                        m.remarks.add(dayNum + "th: Half day due to excess lates");
                        continue;
                    }
                }

                if (minutes >= halfDayMinutes && minutes < fullDayMinutes) {
                    m.halfDays++;
                    m.halfDaysDueToDuration++;
                    m.remarks.add(dayNum + "th: Half day – duration < 8.5 hrs");
                } else if (minutes >= fullDayMinutes) {
                    m.totalFullDays++;
                }

                // Working day OT
                if (inP && outP) {
                    LocalTime inT = LocalTime.parse(in);
                    LocalTime outT = LocalTime.parse(out);
                    long worked = Duration.between(
                            inT.isBefore(shiftStartTime) ? shiftStartTime : inT,
                            outT).toMinutes();
                    if (worked >= fullDayMinutes + otThresholdMinutes) {
                        m.workingDayOTMinutes += (worked - fullDayMinutes);
                    }
                }
            }

            // ---------------- WEEKEND / HOLIDAY ----------------
            else if (isHolidayOrWeekend) {
            	
            	// Weekend/Holiday presence count (Option B: any punch)
            	if (inP || outP) {
            	    m.weekendHolidayPresentDays++;
            	}

                if (punchMiss) {
                    m.totalHalfOTDays++;
                    m.weekendHalfOTMinutes += 240;
                    m.remarks.add(dayNum + "th: Half OT day due to punch miss");
                } else if (inP && outP) {
                    int worked = calculateDuration(in, out);
                    if (worked >= fullOtMinutes) {
                        m.totalOTDays++;
                        m.weekendFullOTMinutes += worked;
                    } else if (worked >= halfDayMinutes) {
                        m.totalHalfOTDays++;
                        m.weekendHalfOTMinutes += worked;
                    }
                }
            }
        }
        return m;
    }

    // ======================= HELPERS =======================

    private int calculateDuration(String in, String out) {
        if (in == null || out == null || in.isEmpty() || out.isEmpty()) return 0;
        long m = Duration.between(LocalTime.parse(in), LocalTime.parse(out)).toMinutes();
        return m < 0 ? (int) (m + 1440) : (int) m;
    }

    private boolean isLate(String in) {
        if (in == null || in.isEmpty()) return false;
        LocalTime t = LocalTime.parse(in);
        return t.isAfter(LocalTime.parse(onTimeEnd))
                && !t.isAfter(LocalTime.parse(halfDayThreshold));
    }

    private boolean isAfterHalfDayThreshold(String in) {
        return in != null && !in.isEmpty()
                && LocalTime.parse(in).isAfter(LocalTime.parse(halfDayThreshold));
    }

    private static boolean isWeekend(int y, int m, int d) {
        DayOfWeek dow = LocalDate.of(y, m, d).getDayOfWeek();
        return dow == DayOfWeek.SATURDAY || dow == DayOfWeek.SUNDAY;
    }

    private double round2(double v) {
        return Math.round(v * 100.0) / 100.0;
    }

    public void exportToCsv(String absolutePath,
            List<AttendanceMerger.EmployeeData> allEmployees,
            int monthDays,
            List<Integer> holidays,
            int year,
            int month) {

try (BufferedWriter writer = new BufferedWriter(new FileWriter(absolutePath))) {

// CSV Header
writer.write(
    "Employee Code," +
    "Total Working Days," +
    "Total Full Days," +
    "Half Days Due To Duration," +
    "Half Days Due To PunchMiss," +
    "Half Days Due To Lates," +
    "Total Half Days," +
    "Total Lates," +
    "Total Absent," +
    "Total Punch Missed," +
    "Total OT Days," +
    "Total Half OT Days," +
    "Total OT Hours"
);
writer.newLine();

// Data rows
for (AttendanceMerger.EmployeeData emp : allEmployees) {

Metrics m = computeMetrics(emp, monthDays, holidays, year, month);

// Weekend/Holiday OT rounding rule:
// exact minutes → exact hours (NO 30-min block rounding)
double otHours = Math.round((m.totalOTMinutes / 60.0) * 100.0) / 100.0;

writer.write(String.format(
        "%s,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%.2f",
        emp.empId,
        m.totalWorkingDays,
        m.totalFullDays,
        m.halfDaysDueToDuration,
        m.halfDaysDueToPunchMiss,
        m.halfDaysDueToLate,
        m.halfDays,
        m.totalLates,
        m.totalAbsent,
        m.totalPunchMissed,
        m.totalOTDays,
        m.totalHalfOTDays,
        otHours
));
writer.newLine();
}

System.out.println("✅ CSV exported successfully to: " + absolutePath);

} catch (Exception e) {
throw new RuntimeException("Error exporting CSV: " + e.getMessage(), e);
}
}



}
