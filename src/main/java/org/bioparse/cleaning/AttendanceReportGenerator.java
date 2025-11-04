package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.*;
import java.util.*;

public class AttendanceReportGenerator {

    private String shiftStart = "09:30";
    private String onTimeEnd = "09:46";        // inclusive end of on-time window
    private String halfDayThreshold = "10:16"; // In-time after 10:16 counts as half day
    private int fullDayMinutes = 510; // 8 hours 30 minutes
    private int halfDayMinutes = 240; // 4 hours
    private int otThresholdMinutes = 30; // Overtime threshold for non-holiday/weekend
    private int fullOtMinutes = 480; // 8 hours for OT days
    private int allowedLatesPerMonth = 5;

    static class Metrics {
        int totalWorkingDays;
        int totalLates;
        int totalAbsent;
        int totalFullDays;
        int halfDays;

        // New breakdown columns for half-day reasons
        int halfDaysDueToLate;       // arrivals after 10:15 OR lates beyond allowed
        int halfDaysDueToPunchMiss;  // single missing punch on working day
        int halfDaysDueToDuration;   // duration >= 240 and < 510

        int totalOTDays;
        int totalOTMinutes;
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
        String[] cols = {
        		 	"Employee Code",
        		    "Total Working Days",
        		    "Total Full Days",
        		    "Half Days Due To Duration",
        		    "Half Days Due To PunchMiss",
        		    "Half Days Due To Lates",
        		    "Total Half Days",
        		    "Total Lates",
        		    "Total Absent",
        		    "Total Punch Missed",
        		    "Total OT Days",
        		    "Total Half OT Days",
        		    "Total OT Hours" 
        };
        for (int i = 0; i < cols.length; i++) {
            header.createCell(i).setCellValue(cols[i]);
        }

        for (AttendanceMerger.EmployeeData emp : allEmployees) {
            Metrics m = computeMetrics(emp, monthDays, holidays, year, month);

            Row row = reportSheet.createRow(rowNum++);
            int c = 0;
            row.createCell(c++).setCellValue(emp.empId);
            row.createCell(c++).setCellValue(m.totalWorkingDays);
            row.createCell(c++).setCellValue(m.totalFullDays);
            row.createCell(c++).setCellValue(m.halfDaysDueToDuration);
            row.createCell(c++).setCellValue(m.halfDaysDueToPunchMiss);
            row.createCell(c++).setCellValue(m.halfDaysDueToLate);
            row.createCell(c++).setCellValue(m.halfDays);
            row.createCell(c++).setCellValue(m.totalLates);
            row.createCell(c++).setCellValue(m.totalAbsent);
            row.createCell(c++).setCellValue(m.totalPunchMissed);
            row.createCell(c++).setCellValue(m.totalOTDays);
            row.createCell(c++).setCellValue(m.totalHalfOTDays);
        
            
            // convert total OT minutes to ceil 30-min blocks and then to hours (0.5 hr per block)
            int blocks = (m.totalOTMinutes + 29) / 30; // ceil to 30-min blocks
            double otHours = blocks * 0.5;
            row.createCell(c++).setCellValue(Math.round(otHours * 100.0) / 100.0); // two-decimal hours
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
        	writer.write("Employee Code,Total Working Days,Total Full Days,Half Days Due To Duration,Half Days Due To PunchMiss,Half Days Due To Extra Lates,Total Half Days,Total Lates,Total Absent,Total Punch Missed,Total OT Days,Total Half OT Days,Total OT Hours\n");
            for (AttendanceMerger.EmployeeData emp : allEmployees) {
                Metrics m = computeMetrics(emp, monthDays, holidays, year, month);
                
                int blocks = (m.totalOTMinutes + 29) / 30; // ceil to 30-min blocks
                double otHours = blocks * 0.5;

                writer.write(String.format("%s,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%.2f\n",
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
                        otHours));
            }
        }
        System.out.println("✅ CSV exported to: " + outputCsvPath);
    }

    private Metrics computeMetrics(AttendanceMerger.EmployeeData emp, int monthDays,
                                   List<Integer> holidays, int year, int month) {

        Metrics m = new Metrics();

        // compute weekend & working days correctly (avoid double-subtracting holidays on weekends)
        int weekendDays = 0;
        for (int day = 1; day <= monthDays; day++) {
            if (isWeekend(year, month, day)) weekendDays++;
        }
        int holidaysOnWorkingDays = 0;
        if (holidays != null) {
            for (Integer hd : holidays) {
                if (hd != null && hd >= 1 && hd <= monthDays && !isWeekend(year, month, hd)) {
                    holidaysOnWorkingDays++;
                }
            }
        }
        m.totalWorkingDays = monthDays - weekendDays - holidaysOnWorkingDays;

        List<String> inTimes = emp.dailyData.getOrDefault("InTime", new ArrayList<>());
        List<String> outTimes = emp.dailyData.getOrDefault("OutTime", new ArrayList<>());
        List<String> status = emp.dailyData.getOrDefault("Status", new ArrayList<>());

        int latesSeen = 0; // counts only working-day lates

        for (int day = 0; day < monthDays; day++) {
            int dayNum = day + 1;
            boolean isHolidayOrWeekend = isWeekend(year, month, dayNum) || (holidays != null && holidays.contains(dayNum));

            String in = AttendanceUtils.safeGet(inTimes, day);
            String out = AttendanceUtils.safeGet(outTimes, day);
            String stat = AttendanceUtils.safeGet(status, day);

            int minutes = calculateDuration(in, out);

            // Absent detection (explicit A)
            if ("A".equalsIgnoreCase(stat) && (in == null || in.isEmpty()) && (out == null || out.isEmpty()) && minutes == 0) {
                m.totalAbsent++;
            }

            // Punch missed: count anytime exactly one punch is missing (applies to any day)
            boolean singlePunchMissing = ((in == null || in.isEmpty()) ^ (out == null || out.isEmpty()));
            if (singlePunchMissing) {
                m.totalPunchMissed++;
            }

            // Only evaluate lates and working-day half-day logic on working days
            if (!isHolidayOrWeekend && !"A".equalsIgnoreCase(stat)) {
                // PRIORITY for half-day classification (only one reason per day):
                // 1) single missing punch -> half-day due to punch miss
                // 2) arrival after 10:15 -> half-day due to late
                // 3) late and latesSeen > allowed -> half-day due to late
                // 4) duration-based half-day (240 <= minutes < 510) -> half-day due to duration

                boolean afterHalfDay = isAfterHalfDayThreshold(in); // in > 10:15
                boolean late = isLate(in);                          // in > 09:45 && in <= 10:15

                // Case 1: single missing punch on working day => half-day due to punch miss
                if (singlePunchMissing) {
                    m.halfDays++;
                    m.halfDaysDueToPunchMiss++;
                    // totalPunchMissed already incremented above
                } else {
                    // Case 2: arrival after 10:15 -> half-day due to late (arrival beyond threshold)
                    if (afterHalfDay) {
                        m.halfDays++;
                        m.halfDaysDueToLate++;
                    } else {
                        // Case 3: late within (09:45,10:15] and exceeds monthly allowed -> half-day due to late
                        if (late) {
                            latesSeen++;
                            m.totalLates++;
                            if (latesSeen > allowedLatesPerMonth) {
                                m.halfDays++;
                                m.halfDaysDueToLate++;
                            }
                        }
                        // Case 4: duration-based half day (only if not already half-day by above)
                        if (minutes >= halfDayMinutes && minutes < fullDayMinutes) {
                            m.halfDays++;
                            m.halfDaysDueToDuration++;
                        }
                        // If minutes >= fullDayMinutes and not afterHalfDay => full day
                        if (minutes >= fullDayMinutes && !afterHalfDay) {
                            m.totalFullDays++;
                        }
                    }
                }

                // Working-day OT minutes accumulation 
                if (minutes > fullDayMinutes + otThresholdMinutes) {
                    int otMinutes = minutes - fullDayMinutes;
                    m.totalOTMinutes += otMinutes;
                }
            }
            // Holiday/weekend: OT and punch-miss logic (punch-miss counted above)
            else if (isHolidayOrWeekend) {
                boolean inPresent = in != null && !in.isEmpty();
                boolean outPresent = out != null && !out.isEmpty();

                if (inPresent ^ outPresent) {
                    // one punch missing on weekend => half OT day (and punch missed already counted)
                    m.totalHalfOTDays++;
                } else if (inPresent && outPresent) {
                	if (minutes >= fullOtMinutes) {
                	    m.totalOTDays++;
                	    int otMinutes = minutes - fullOtMinutes;
                	    m.totalOTMinutes += otMinutes;
                	} else if (minutes >= halfDayMinutes) {
                	    m.totalHalfOTDays++;
                	    m.totalOTMinutes += minutes;
                	} else if (minutes > 0) {
                	    m.totalOTMinutes += minutes;
                	}
                }
            }
        }

        return m;
    }

    private int calculateDuration(String inTime, String outTime) {
        if (inTime == null || inTime.isEmpty() || outTime == null || outTime.isEmpty()) {
            return 0;
        }
        try {
            LocalTime in = LocalTime.parse(inTime);
            LocalTime out = LocalTime.parse(outTime);
            // Handle cases where outTime is on the next day (e.g., after midnight)
            long minutes = Duration.between(in, out).toMinutes();
            if (minutes < 0) {
                minutes += 24 * 60; // Add 24 hours if negative
            }
            return (int) minutes;
        } catch (Exception e) {
            return 0;
        }
    }

    // Late: strictly after 09:45 and up to 10:15 (inclusive)
    private boolean isLate(String inTime) {
        if (inTime == null || inTime.isEmpty()) return false;
        try {
            LocalTime in = LocalTime.parse(inTime);
            LocalTime onEnd = LocalTime.parse(onTimeEnd);          // 09:45
            LocalTime halfDayThresh = LocalTime.parse(halfDayThreshold); // 10:15
            return in.isAfter(onEnd) && !in.isAfter(halfDayThresh); // (09:45, 10:15]
        } catch (Exception e) {
            return false;
        }
    }

    // After 10:15 => half day
    private boolean isAfterHalfDayThreshold(String inTime) {
        if (inTime == null || inTime.isEmpty()) return false;
        try {
            LocalTime in = LocalTime.parse(inTime);
            LocalTime halfDayThresh = LocalTime.parse(halfDayThreshold);
            return in.isAfter(halfDayThresh); // strictly > 10:15
        } catch (Exception e) {
            return false;
        }
    }

    private static boolean isWeekend(int year, int month, int day) {
        try {
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
