package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class AttendanceQueryViewer {

    private Workbook wb;
    private Sheet masterSheet;
    private List<AttendanceMerger.EmployeeData> employees;

    public AttendanceQueryViewer(String mergedFile, List<AttendanceMerger.EmployeeData> employees) throws Exception {
        FileInputStream fis = new FileInputStream(mergedFile);
        wb = new XSSFWorkbook(fis);
        masterSheet = wb.getSheet("Master");
        this.employees = employees;
    }

    public void showByEmployee(String empCode) {
        AttendanceMerger.EmployeeData emp = employees.stream()
                .filter(e -> e.empId.equalsIgnoreCase(empCode))
                .findFirst()
                .orElse(null);

        if (emp == null) {
            JOptionPane.showMessageDialog(null, "Employee not found: " + empCode);
            return;
        }

        showEmployeeTable(emp);
    }

    public void showByDate(int day) {
        if (day <= 0) return;

        String[] columns = {"Employee Code", "Employee Name", "InTime", "OutTime", "Duration", "Status"};
        DefaultTableModel model = new DefaultTableModel(columns, 0);

        for (AttendanceMerger.EmployeeData emp : employees) {
            String in = AttendanceUtils.safeGet(emp.dailyData.get("InTime"), day - 1);
            String out = AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), day - 1);
            String dur = AttendanceUtils.safeGet(emp.dailyData.get("Duration"), day - 1);
            String stat = AttendanceUtils.safeGet(emp.dailyData.get("Status"), day - 1);

            model.addRow(new Object[]{emp.empId, emp.empName, in, out, dur, stat});
        }

        showTable("Attendance for Day " + day, model);
    }

    public void showByDateRange(int startDay, int endDay) {
        if (startDay > endDay || startDay <= 0) return;

        String[] columns = {"Employee Code", "Employee Name", "Day", "InTime", "OutTime", "Duration", "Status"};
        DefaultTableModel model = new DefaultTableModel(columns, 0);

        for (AttendanceMerger.EmployeeData emp : employees) {
            for (int d = startDay; d <= endDay; d++) {
                String in = AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d - 1);
                String out = AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), d - 1);
                String dur = AttendanceUtils.safeGet(emp.dailyData.get("Duration"), d - 1);
                String stat = AttendanceUtils.safeGet(emp.dailyData.get("Status"), d - 1);

                model.addRow(new Object[]{emp.empId, emp.empName, d, in, out, dur, stat});
            }
        }

        showTable("Attendance from Day " + startDay + " to " + endDay, model);
    }

    private void showEmployeeTable(AttendanceMerger.EmployeeData emp) {
        String[] columns = {"Day", "InTime", "OutTime", "Duration", "Status"};
        DefaultTableModel model = new DefaultTableModel(columns, 0);

        int monthDays = emp.dailyData.values().stream().findFirst().orElse(new ArrayList<>()).size();
        for (int d = 0; d < monthDays; d++) {
            String in = AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d);
            String out = AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), d);
            String dur = AttendanceUtils.safeGet(emp.dailyData.get("Duration"), d);
            String stat = AttendanceUtils.safeGet(emp.dailyData.get("Status"), d);
            model.addRow(new Object[]{d + 1, in, out, dur, stat});
        }

        showTable("Attendance for " + emp.empId + " : " + emp.empName, model);
    }

    private void showTable(String title, DefaultTableModel model) {
        JTable table = new JTable(model);
        JScrollPane scrollPane = new JScrollPane(table);
        table.setFillsViewportHeight(true);

        JFrame frame = new JFrame(title);
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setSize(800, 600);
        frame.setLayout(new BorderLayout());
        frame.add(scrollPane, BorderLayout.CENTER);
        frame.setVisible(true);
    }
}