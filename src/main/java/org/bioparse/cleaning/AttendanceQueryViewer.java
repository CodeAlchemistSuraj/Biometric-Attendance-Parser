package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import javax.swing.table.TableRowSorter;
import java.awt.*;
import java.awt.Color;
import java.io.*;
import java.text.DateFormatSymbols;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class AttendanceQueryViewer {

    private Workbook wb;
    private Sheet masterSheet;
    private List<AttendanceMerger.EmployeeData> employees;
    private JTable currentTable;
    private String currentTitle;
    private TableRowSorter<DefaultTableModel> sorter;
    private JTextField searchField;
    private JComboBox<String> statusFilter;
    private JComboBox<String> sortColumn;
    private JCheckBox sortDescending;

    // Store the processed month/year (passed from caller or detected)
    private int currentYear;
    private int currentMonth;

    // PRIMARY CONSTRUCTOR: requires year and month
    public AttendanceQueryViewer(String mergedFile, List<AttendanceMerger.EmployeeData> employees, int year, int month) throws Exception {
        FileInputStream fis = new FileInputStream(mergedFile);
        wb = new XSSFWorkbook(fis);
        masterSheet = wb.getSheet("Master");
        this.employees = employees;
        this.currentYear = year;
        this.currentMonth = month;
    }

    // IMPROVED OVERLOAD: attempts to detect month/year from workbook; falls back to system month/year
    public AttendanceQueryViewer(String mergedFile, List<AttendanceMerger.EmployeeData> employees) throws Exception {
        FileInputStream fis = new FileInputStream(mergedFile);
        wb = new XSSFWorkbook(fis);
        masterSheet = wb.getSheet("Master");
        this.employees = employees;

        // Try detect month/year from workbook; fallback to current month/year
        int detectedYear = -1;
        int detectedMonth = -1;
        try {
            int[] y_m = detectMonthYearFromWorkbook(masterSheet);
            detectedYear = y_m[0];
            detectedMonth = y_m[1];
        } catch (Exception ex) {
            // ignore detection error, fallback below
        }

        if (detectedYear > 0 && detectedMonth > 0) {
            this.currentYear = detectedYear;
            this.currentMonth = detectedMonth;
        } else {
            // fallback: use current system date
            this.currentYear = LocalDate.now().getYear();
            this.currentMonth = LocalDate.now().getMonthValue();
            System.err.println("Warning: Could not detect processed month/year from workbook. Falling back to current system month/year: "
                    + this.currentMonth + "/" + this.currentYear);
        }
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
        // Use the stored year/month from processed data
        showByDate(day, currentYear, currentMonth);
    }

    public void showByDate(int day, int year, int month) {
        if (day <= 0 || day > YearMonth.of(year, month).lengthOfMonth()) return;

        String[] columns = {"Employee Code", "Employee Name", "InTime", "OutTime", "Duration", "Status"};
        DefaultTableModel model = createTableModel(columns);

        LocalDate date = LocalDate.of(year, month, day);
        String formattedDate = date.format(DateTimeFormatter.ofPattern("EEE, MMM d"));

        for (AttendanceMerger.EmployeeData emp : employees) {
            String in = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("InTime"), day - 1));
            String out = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), day - 1));
            String dur = calculateDuration(in, out);
            String stat = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Status"), day - 1));

            model.addRow(new Object[]{emp.empId, emp.empName, in, out, dur, stat});
        }

        showTable("Attendance for " + formattedDate, model, columns);
    }

    public void showByDateRange(int startDay, int endDay, int year, int month) {
        if (startDay > endDay || startDay <= 0 || endDay > YearMonth.of(year, month).lengthOfMonth()) return;

        String[] columns = {"Employee Code", "Employee Name", "Day", "InTime", "OutTime", "Duration", "Status"};
        DefaultTableModel model = createTableModel(columns);

        for (AttendanceMerger.EmployeeData emp : employees) {
            for (int d = startDay; d <= endDay; d++) {
                LocalDate date = LocalDate.of(year, month, d);
                String formattedDate = date.format(DateTimeFormatter.ofPattern("EEE, MMM d"));
                String in = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d - 1));
                String out = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), d - 1));
                String dur = calculateDuration(in, out);
                String stat = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Status"), d - 1));

                model.addRow(new Object[]{emp.empId, emp.empName, formattedDate, in, out, dur, stat});
            }
        }

        showTable("Attendance from Day " + startDay + " to " + endDay, model, columns);
    }

    // FIXED: Use dynamic month/year + validation
    private void showEmployeeTable(AttendanceMerger.EmployeeData emp) {
        String[] columns = {"Day", "InTime", "OutTime", "Duration", "Status"};
        DefaultTableModel model = createTableModel(columns);

        // Get actual days in the month
        YearMonth ym = YearMonth.of(currentYear, currentMonth);
        int maxDays = ym.lengthOfMonth();  // e.g., 31 for Oct, 30 for Sep

        // Choose a sample list from dailyData (prefer Status, otherwise any first value)
        List<String> sample = emp.dailyData.getOrDefault("Status",
                emp.dailyData.values().stream().findFirst().orElse(new ArrayList<>()));
        int monthDays = Math.min(sample.size(), maxDays);

        for (int d = 0; d < monthDays; d++) {
            int day = d + 1;
            try {
                LocalDate date = LocalDate.of(currentYear, currentMonth, day);
                String formattedDate = date.format(DateTimeFormatter.ofPattern("EEE, MMM d"));
                String in = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d));
                String out = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), d));
                String dur = calculateDuration(in, out);
                String stat = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Status"), d));
                model.addRow(new Object[]{formattedDate, in, out, dur, stat});
            } catch (DateTimeException e) {
                // Graceful skip for invalid dates
                System.err.println("Skipping invalid date: " + currentYear + "-" + currentMonth + "-" + day + " for " + emp.empId);
                model.addRow(new Object[]{"Invalid Date", "-", "-", "-", "Error"});
            }
        }

        String monthName = new DateFormatSymbols().getMonths()[currentMonth - 1];  // e.g., "October"
        showTable("Attendance for " + monthName + " " + currentYear + " - " + emp.empId + " : " + emp.empName, model, columns);
    }

    private String calculateDuration(String inTime, String outTime) {
        if (inTime == null || inTime.isEmpty() || outTime == null || outTime.isEmpty()) {
            return "-";
        }
        try {
            // Accept HH:mm, HH:mm:ss or attempt to parse flexibly
            LocalTime in = parseLocalTimeFlexible(inTime);
            LocalTime out = parseLocalTimeFlexible(outTime);
            long minutes = Duration.between(in, out).toMinutes();
            if (minutes < 0) {
                minutes += 24 * 60; // Handle next-day outTime
            }
            long hours = minutes / 60;
            long mins = minutes % 60;
            return String.format("%d:%02d", hours, mins);
        } catch (Exception e) {
            return "-";
        }
    }

    private LocalTime parseLocalTimeFlexible(String t) {
        // Try parsing common formats
        try {
            return LocalTime.parse(t);
        } catch (DateTimeParseException ignored) {}
        try {
            DateTimeFormatter f = DateTimeFormatter.ofPattern("H:mm");
            return LocalTime.parse(t, f);
        } catch (DateTimeParseException ignored) {}
        try {
            DateTimeFormatter f2 = DateTimeFormatter.ofPattern("hh:mm a");
            return LocalTime.parse(t, f2);
        } catch (DateTimeParseException ignored) {}
        // fallback
        throw new DateTimeParseException("Unsupported time format", t, 0);
    }

    private DefaultTableModel createTableModel(String[] columns) {
        return new DefaultTableModel(columns, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // Read-only
            }
        };
    }

    private String displayValue(String s) {
        return (s == null || s.trim().isEmpty()) ? "-" : s;
    }

    private void showTable(String title, DefaultTableModel model, String[] columns) {
        currentTable = new JTable(model);
        currentTitle = title;
        sorter = new TableRowSorter<>(model);
        currentTable.setRowSorter(sorter);

        // Styling
        currentTable.setShowGrid(true);
        currentTable.setGridColor(Color.LIGHT_GRAY);
        currentTable.setRowHeight(22);
        currentTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        // Custom renderer for highlighting and centering (safe checks)
        DefaultTableCellRenderer cellRenderer = new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
                                                          boolean hasFocus, int row, int column) {
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                setHorizontalAlignment(SwingConstants.CENTER);
                String text = (value == null || value.toString().trim().isEmpty()) ? "-" : value.toString();
                setText(text);

                int statusCol = -1;
                for (int i = 0; i < table.getColumnCount(); i++) {
                    if (table.getColumnName(i).equals("Status")) {
                        statusCol = i;
                        break;
                    }
                }
                String status = "";
                if (statusCol != -1) {
                    Object sObj = table.getValueAt(row, statusCol);
                    status = (sObj == null) ? "" : sObj.toString();
                }

                if ("A".equals(status)) {
                    c.setBackground(new Color(255, 204, 204)); // Light red
                } else if ("H".equals(status) || "WO".equals(status) || "WOP".equals(status)) {
                    c.setBackground(new Color(220, 220, 220)); // Light gray
                } else {
                    c.setBackground(row % 2 == 0 ? Color.WHITE : new Color(240, 240, 240)); // Alternating rows
                }
                return c;
            }
        };
        currentTable.setDefaultRenderer(Object.class, cellRenderer);

        // Set column widths
        if (currentTable.getColumnModel().getColumnCount() >= 2) {
            try {
                currentTable.getColumnModel().getColumn(0).setPreferredWidth(100);
                if (currentTable.getColumnModel().getColumnCount() > 1) {
                    currentTable.getColumnModel().getColumn(1).setPreferredWidth(200);
                }
            } catch (Exception ignored) {}
        }

        JScrollPane scrollPane = new JScrollPane(currentTable);

        // Toolbar
        JPanel toolbar = new JPanel(new FlowLayout(FlowLayout.LEFT));
        searchField = new JTextField(15);
        searchField.setToolTipText("Search by Employee Code or Name");
        searchField.addActionListener(e -> applyFilters());

        statusFilter = new JComboBox<>(new String[]{"All", "P", "A", "H", "WO", "WOP"});
        statusFilter.addActionListener(e -> applyFilters());

        sortColumn = new JComboBox<>(columns);
        sortDescending = new JCheckBox("Descending");
        JButton sortButton = new JButton("Sort");
        sortButton.setToolTipText("Sort table by selected column");
        sortButton.addActionListener(e -> applySort());

        JButton clearFilters = new JButton("Clear Filters");
        clearFilters.setToolTipText("Reset all filters");
        clearFilters.addActionListener(e -> clearFilters());

        JButton exportExcel = new JButton("Export to Excel");
        exportExcel.setToolTipText("Export table to Excel file");
        exportExcel.addActionListener(e -> exportTableToExcel());

        JButton exportCsv = new JButton("Export to CSV");
        exportCsv.setToolTipText("Export table to CSV file");
        exportCsv.addActionListener(e -> exportTableToCsv());

        toolbar.add(new JLabel("Search:"));
        toolbar.add(searchField);
        toolbar.add(new JLabel("Status:"));
        toolbar.add(statusFilter);
        toolbar.add(new JLabel("Sort by:"));
        toolbar.add(sortColumn);
        toolbar.add(sortDescending);
        toolbar.add(sortButton);
        toolbar.add(clearFilters);
        toolbar.add(exportExcel);
        toolbar.add(exportCsv);

        JFrame frame = new JFrame(title);
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setSize(1000, 700);
        frame.setLayout(new BorderLayout());
        frame.add(toolbar, BorderLayout.NORTH);
        frame.add(scrollPane, BorderLayout.CENTER);
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }

    private void applyFilters() {
        RowFilter<DefaultTableModel, Object> filter = RowFilter.andFilter(Arrays.asList(
                createSearchFilter(),
                createStatusFilter()
        ));
        sorter.setRowFilter(filter);
    }

    private RowFilter<DefaultTableModel, Object> createSearchFilter() {
        String searchText = searchField.getText().trim();
        if (searchText.isEmpty()) {
            return RowFilter.regexFilter(".*");
        }
        try {
            return RowFilter.regexFilter("(?i)" + Pattern.quote(searchText), 0, 1); // Search in Employee Code and Name
        } catch (Exception ex) {
            return RowFilter.regexFilter(".*");
        }
    }

    private RowFilter<DefaultTableModel, Object> createStatusFilter() {
        String selectedStatus = (String) statusFilter.getSelectedItem();
        if (selectedStatus == null || selectedStatus.equals("All")) {
            return RowFilter.regexFilter(".*");
        }
        int statusCol = -1;
        for (int i = 0; i < currentTable.getColumnCount(); i++) {
            if (currentTable.getColumnName(i).equals("Status")) {
                statusCol = i;
                break;
            }
        }
        if (statusCol == -1) return RowFilter.regexFilter(".*");
        return RowFilter.regexFilter("^" + Pattern.quote(selectedStatus) + "$", statusCol);
    }

    private void applySort() {
        String selectedColumn = (String) sortColumn.getSelectedItem();
        boolean descending = sortDescending.isSelected();
        int columnIndex = -1;
        for (int i = 0; i < currentTable.getColumnCount(); i++) {
            if (currentTable.getColumnName(i).equals(selectedColumn)) {
                columnIndex = i;
                break;
            }
        }
        if (columnIndex != -1) {
            sorter.setSortKeys(Collections.singletonList(
                    new RowSorter.SortKey(columnIndex, descending ? SortOrder.DESCENDING : SortOrder.ASCENDING)
            ));
        }
    }

    private void clearFilters() {
        searchField.setText("");
        statusFilter.setSelectedItem("All");
        sorter.setRowFilter(null);
        sorter.setSortKeys(null);
    }

    private void exportTableToExcel() {
        if (currentTable == null) {
            JOptionPane.showMessageDialog(null, "No table data to export!");
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Excel File");
        fileChooser.setSelectedFile(new File(currentTitle + ".xlsx"));

        if (fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            if (!fileToSave.getName().toLowerCase().endsWith(".xlsx")) {
                fileToSave = new File(fileToSave.getAbsolutePath() + ".xlsx");
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

    private void exportTableToCsv() {
        if (currentTable == null) {
            JOptionPane.showMessageDialog(null, "No table data to export!");
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save CSV File");
        fileChooser.setSelectedFile(new File(currentTitle + ".csv"));

        if (fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            if (!fileToSave.getName().toLowerCase().endsWith(".csv")) {
                fileToSave = new File(fileToSave.getAbsolutePath() + ".csv");
            }
            try {
                exportTableDataToCsv(fileToSave);
                JOptionPane.showMessageDialog(null, "Data exported successfully to:\n" + fileToSave.getAbsolutePath());
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(null, "Error exporting data: " + ex.getMessage(), "Export Error", JOptionPane.ERROR_MESSAGE);
                ex.printStackTrace();
            }
        }
    }

    private void exportTableDataToExcel(File outputFile) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Attendance Data");

        TableModel model = currentTable.getModel();
        int rowCount = model.getRowCount();
        int colCount = model.getColumnCount();

        // Header style
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

        // Data rows
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setWrapText(true);

        for (int row = 0; row < rowCount; row++) {
            Row dataRow = sheet.createRow(row + 1);
            for (int col = 0; col < colCount; col++) {
                Cell cell = dataRow.createCell(col);
                Object value = model.getValueAt(row, col);
                String cellValue = (value == null) ? "" : value.toString();
                if ("-".equals(cellValue)) {
                    cellValue = "";
                }
                cell.setCellValue(cellValue);
                cell.setCellStyle(dataStyle);
            }
        }

        for (int col = 0; col < colCount; col++) {
            sheet.autoSizeColumn(col);
        }

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }
        workbook.close();
    }

    private void exportTableDataToCsv(File outputFile) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile))) {
            TableModel model = currentTable.getModel();
            int colCount = model.getColumnCount();
            int rowCount = model.getRowCount();

            // Write header
            for (int col = 0; col < colCount; col++) {
                writer.write("\"" + model.getColumnName(col).replace("\"", "\"\"") + "\"");
                if (col < colCount - 1) writer.write(",");
            }
            writer.newLine();

            // Write data
            for (int row = 0; row < rowCount; row++) {
                for (int col = 0; col < colCount; col++) {
                    Object value = model.getValueAt(row, col);
                    String cellValue = (value == null || "-".equals(value.toString())) ? "" : value.toString();
                    writer.write("\"" + cellValue.replace("\"", "\"\"") + "\"");
                    if (col < colCount - 1) writer.write(",");
                }
                writer.newLine();
            }
        }
    }

    // --- Workbook detection helpers ---
    private int[] detectMonthYearFromWorkbook(Sheet sheet) {
        if (sheet == null) throw new IllegalArgumentException("sheet is null");

     // Regex to capture month name (short or long) and year
        String monthNames = "Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?";
        Pattern p = Pattern.compile("(?i)\\b(" + monthNames + ")\\b[\\s\\-\\/,]*?(\\d{4})");

        // search a small area of sheet
        int maxRows = Math.min(40, sheet.getLastRowNum() + 1);
        for (int r = 0; r < maxRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            int maxCols = Math.min(10, row.getLastCellNum() == -1 ? 10 : row.getLastCellNum());
            for (int c = 0; c < maxCols; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) continue;
                String text = getCellAsString(cell);
                if (text == null || text.trim().isEmpty()) continue;
                Matcher m = p.matcher(text);
                if (m.find()) {
                    String monthText = m.group(1);
                    String yearText = m.group(2);
                    int month = parseMonthNameToNumber(monthText);
                    int year = Integer.parseInt(yearText);
                    if (month >= 1 && year > 0) return new int[]{year, month};
                }
            }
        }

        // also check sheet name
        String sheetName = sheet.getSheetName();
        if (sheetName != null) {
            Matcher m2 = p.matcher(sheetName);
            if (m2.find()) {
                int month = parseMonthNameToNumber(m2.group(1));
                int year = Integer.parseInt(m2.group(2));
                return new int[]{year, month};
            }
        }

        throw new RuntimeException("Processed month/year not found in sheet.");
    }

    private String getCellAsString(Cell cell) {
        try {
            if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue();
            if (cell.getCellType() == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date dt = cell.getDateCellValue();
                    return new java.text.SimpleDateFormat("MMMM yyyy").format(dt);
                } else {
                    return Double.toString(cell.getNumericCellValue());
                }
            }
            return cell.toString();
        } catch (Exception e) {
            return null;
        }
    }

    private int parseMonthNameToNumber(String mname) {
        if (mname == null) return -1;
        mname = mname.toLowerCase();
        switch (mname) {
            case "jan": case "january": return 1;
            case "feb": case "february": return 2;
            case "mar": case "march": return 3;
            case "apr": case "april": return 4;
            case "may": return 5;
            case "jun": case "june": return 6;
            case "jul": case "july": return 7;
            case "aug": case "august": return 8;
            case "sep": case "september": return 9;
            case "oct": case "october": return 10;
            case "nov": case "november": return 11;
            case "dec": case "december": return 12;
            default: return -1;
        }
    }
}
