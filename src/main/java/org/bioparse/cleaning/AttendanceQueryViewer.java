package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.awt.*;
import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class AttendanceQueryViewer {

    private Workbook wb;
    private Sheet masterSheet;
    private List<AttendanceMerger.EmployeeData> employees;
    private JTable currentTable;
    private String currentTitle;

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
        DefaultTableModel model = createTableModel(columns);

        for (AttendanceMerger.EmployeeData emp : employees) {
            String in = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("InTime"), day - 1));
            String out = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), day - 1));
            String dur = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Duration"), day - 1));
            String stat = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Status"), day - 1));

            model.addRow(new Object[]{emp.empId, emp.empName, in, out, dur, stat});
        }

        showTable("Attendance for Day " + day, model);
    }

    public void showByDateRange(int startDay, int endDay) {
        if (startDay > endDay || startDay <= 0) return;

        String[] columns = {"Employee Code", "Employee Name", "Day", "InTime", "OutTime", "Duration", "Status"};
        DefaultTableModel model = createTableModel(columns);

        for (AttendanceMerger.EmployeeData emp : employees) {
            for (int d = startDay; d <= endDay; d++) {
                String in = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d - 1));
                String out = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), d - 1));
                String dur = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Duration"), d - 1));
                String stat = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Status"), d - 1));

                model.addRow(new Object[]{emp.empId, emp.empName, d, in, out, dur, stat});
            }
        }

        showTable("Attendance from Day " + startDay + " to " + endDay, model);
    }

    private void showEmployeeTable(AttendanceMerger.EmployeeData emp) {
        String[] columns = {"Day", "InTime", "OutTime", "Duration", "Status"};
        DefaultTableModel model = createTableModel(columns);

        int monthDays = emp.dailyData.values().stream().findFirst().orElse(new ArrayList<>()).size();
        for (int d = 0; d < monthDays; d++) {
            String in = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d));
            String out = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), d));
            String dur = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Duration"), d));
            String stat = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Status"), d));
            model.addRow(new Object[]{d + 1, in, out, dur, stat});
        }

        showTable("Attendance for " + emp.empId + " : " + emp.empName, model);
    }

    // Create a non-editable table model
    private DefaultTableModel createTableModel(String[] columns) {
        return new DefaultTableModel(columns, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // make table read-only
            }
        };
    }

    // Replace null/empty strings with "-" for display
    private String displayValue(String s) {
        return (s == null || s.trim().isEmpty()) ? "-" : s;
    }

    private void showTable(String title, DefaultTableModel model) {
        JTable table = new JTable(model);
        this.currentTable = table;
        this.currentTitle = title;

        // Show grid lines and a subtle color
        table.setShowGrid(true);
        table.setGridColor(Color.LIGHT_GRAY);

        // Use a renderer that displays "-" for null/empty and centers text
        DefaultTableCellRenderer cellRenderer = new DefaultTableCellRenderer() {
            @Override
            public void setValue(Object value) {
                String text = (value == null || value.toString().trim().isEmpty()) ? "-" : value.toString();
                setText(text);
            }
        };
        cellRenderer.setHorizontalAlignment(SwingConstants.CENTER);
        table.setDefaultRenderer(Object.class, cellRenderer);

        // Header settings
        table.getTableHeader().setReorderingAllowed(false);
        table.setAutoCreateRowSorter(true); // enable sorting by clicking headers
        table.setRowHeight(22);
        table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        // Optional: make first two columns a little wider (Employee code/name)
        if (table.getColumnModel().getColumnCount() >= 2) {
            try {
                table.getColumnModel().getColumn(0).setPreferredWidth(100);
                table.getColumnModel().getColumn(1).setPreferredWidth(200);
            } catch (Exception ignored) {}
        }

        JScrollPane scrollPane = new JScrollPane(table);

        // Create export button
        JButton exportButton = new JButton("Export to Excel");
        exportButton.addActionListener(e -> exportTableToExcel());
        
        // Create button panel
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(exportButton);

        JFrame frame = new JFrame(title);
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setSize(900, 600);
        frame.setLayout(new BorderLayout());
        frame.add(scrollPane, BorderLayout.CENTER);
        frame.add(buttonPanel, BorderLayout.SOUTH);

        // center frame on screen
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }

    /**
     * Export the current table data to Excel file
     */
    private void exportTableToExcel() {
        if (currentTable == null) {
            JOptionPane.showMessageDialog(null, "No table data to export!");
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Excel File");
        fileChooser.setSelectedFile(new java.io.File(currentTitle + ".xlsx"));
        
        int userSelection = fileChooser.showSaveDialog(null);
        
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            java.io.File fileToSave = fileChooser.getSelectedFile();
            
            // Ensure .xlsx extension
            if (!fileToSave.getName().toLowerCase().endsWith(".xlsx")) {
                fileToSave = new java.io.File(fileToSave.getAbsolutePath() + ".xlsx");
            }
            
            try {
                exportTableDataToExcel(fileToSave);
                JOptionPane.showMessageDialog(null, "Data exported successfully to:\n" + fileToSave.getAbsolutePath());
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(null, "Error exporting data: " + ex.getMessage(), "Export Error", JOptionPane.ERROR_MESSAGE);
                ex.printStackTrace();
            }
        }
    }

    /**
     * Actual Excel export implementation
     */
    private void exportTableDataToExcel(java.io.File outputFile) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Attendance Data");
        
        TableModel model = currentTable.getModel();
        int rowCount = model.getRowCount();
        int colCount = model.getColumnCount();
        
        // Create header row with style
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        
        Row headerRow = sheet.createRow(0);
        for (int col = 0; col < colCount; col++) {
            Cell cell = headerRow.createCell(col);
            cell.setCellValue(model.getColumnName(col));
            cell.setCellStyle(headerStyle);
        }
        
        // Create data rows
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setWrapText(true);
        
        for (int row = 0; row < rowCount; row++) {
            Row dataRow = sheet.createRow(row + 1);
            for (int col = 0; col < colCount; col++) {
                Cell cell = dataRow.createCell(col);
                Object value = model.getValueAt(row, col);
                String cellValue = (value == null) ? "" : value.toString();
                
                // Replace display "-" with empty string for export
                if ("-".equals(cellValue)) {
                    cellValue = "";
                }
                
                cell.setCellValue(cellValue);
                cell.setCellStyle(dataStyle);
            }
        }
        
        // Auto-size columns
        for (int col = 0; col < colCount; col++) {
            sheet.autoSizeColumn(col);
        }
        
        // Write to file
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }
        
        workbook.close();
    }
}