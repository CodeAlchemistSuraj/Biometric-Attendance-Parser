package org.bioparse.cleaning;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import com.formdev.flatlaf.FlatLightLaf; // Add dependency: com.formdev:flatlaf

public class AttendanceAppGUI extends JFrame {

    private CardLayout cardLayout;
    private JPanel mainPanel, sidebarPanel;
    private JTextField inputFileField, mergedFileField, reportFileField, holidaysField, yearField, monthField;
    private JButton selectInputBtn, selectMergedBtn, selectReportBtn, processBtn, exportCsvBtn;
    private JTextField empCodeField, fromDayField, toDayField;
    private JButton queryBtn;
    private JTable dataTable, reportTable;
    private DefaultTableModel dataModel, reportModel;
    private AttendanceMerger.MergeResult mergeResult;
    private AttendanceReportGenerator reportGenerator = new AttendanceReportGenerator();
    private AttendanceQueryViewer queryViewer;

    public AttendanceAppGUI() {
        try {
            UIManager.setLookAndFeel(new FlatLightLaf()); // Modern look & feel
        } catch (Exception e) {
            e.printStackTrace();
        }

        setTitle("Biometric Attendance Parser");
        setSize(1200, 800);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // Menu Bar
        JMenuBar menuBar = new JMenuBar();
        JMenu fileMenu = new JMenu("File");
        fileMenu.add(new JMenuItem("Open Input..."));
        fileMenu.add(new JMenuItem("Save Report..."));
        JMenu helpMenu = new JMenu("Help");
        helpMenu.add(new JMenuItem("About"));
        menuBar.add(fileMenu);
        menuBar.add(helpMenu);
        setJMenuBar(menuBar);

        // Sidebar Navigation
        sidebarPanel = new JPanel(new GridLayout(4, 1, 0, 10));
        sidebarPanel.setPreferredSize(new Dimension(200, 0));
        sidebarPanel.setBackground(new Color(240, 240, 240));
        addButtonToSidebar("Home", e -> showPanel("HOME"));
        addButtonToSidebar("Configuration", e -> showPanel("CONFIG"));
        addButtonToSidebar("Data Viewer", e -> showPanel("DATA"));
        addButtonToSidebar("Report Viewer", e -> showPanel("REPORT"));
        add(sidebarPanel, BorderLayout.WEST);

        // Main Content with CardLayout
        cardLayout = new CardLayout();
        mainPanel = new JPanel(cardLayout);
        mainPanel.add(createHomePanel(), "HOME");
        mainPanel.add(createConfigWizardPanel(), "CONFIG");
        mainPanel.add(createDataViewerPanel(), "DATA");
        mainPanel.add(createReportViewerPanel(), "REPORT");
        add(mainPanel, BorderLayout.CENTER);

        showPanel("HOME");
    }

    private void addButtonToSidebar(String text, java.awt.event.ActionListener listener) {
        JButton button = new JButton(text);
        button.addActionListener(listener);
        button.setFont(new Font("Arial", Font.BOLD, 14));
        button.setBackground(new Color(100, 181, 246));
        button.setForeground(Color.WHITE);
        button.setToolTipText("Go to " + text);
        sidebarPanel.add(button);
    }

    private void showPanel(String panelName) {
        cardLayout.show(mainPanel, panelName);
    }

    private JPanel createHomePanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(20, 0, 0, 0);
        gbc.anchor = GridBagConstraints.CENTER;
        panel.add(new JLabel("Welcome to Attendance Parser Dashboard"), gbc);
        // TODO: Add summary cards (e.g., recent files, stats)
        return panel;
    }

    private JPanel createConfigWizardPanel() {
        JPanel wizardPanel = new JPanel(new CardLayout());
        CardLayout wizardLayout = (CardLayout) wizardPanel.getLayout();

        // Step 1: File Selection
        JPanel step1 = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;

        gbc.gridx = 0; gbc.gridy = 0;
        step1.add(new JLabel("Input Excel File:"), gbc);
        gbc.gridy++;
        inputFileField = new JTextField();
        selectInputBtn = new JButton("Select");
        selectInputBtn.addActionListener(this::selectFile);
        JPanel inputPanel = new JPanel(new BorderLayout());
        inputPanel.add(inputFileField, BorderLayout.CENTER);
        inputPanel.add(selectInputBtn, BorderLayout.EAST);
        step1.add(inputPanel, gbc);

        gbc.gridy++;
        step1.add(new JLabel("Merged Output File:"), gbc);
        gbc.gridy++;
        mergedFileField = new JTextField("Merged.xlsx");
        selectMergedBtn = new JButton("Select");
        selectMergedBtn.addActionListener(this::selectFile);
        JPanel mergedPanel = new JPanel(new BorderLayout());
        mergedPanel.add(mergedFileField, BorderLayout.CENTER);
        mergedPanel.add(selectMergedBtn, BorderLayout.EAST);
        step1.add(mergedPanel, gbc);

        gbc.gridy++;
        step1.add(new JLabel("Report Output File:"), gbc);
        gbc.gridy++;
        reportFileField = new JTextField("Employee_Report.xlsx");
        selectReportBtn = new JButton("Select");
        selectReportBtn.addActionListener(this::selectFile);
        JPanel reportPanel = new JPanel(new BorderLayout());
        reportPanel.add(reportFileField, BorderLayout.CENTER);
        reportPanel.add(selectReportBtn, BorderLayout.EAST);
        step1.add(reportPanel, gbc);

        gbc.gridy++;
        JButton nextBtn = new JButton("Next");
        nextBtn.addActionListener(e -> wizardLayout.next(wizardPanel));
        step1.add(nextBtn, gbc);

        wizardPanel.add(step1, "STEP1");

        // Step 2: Holidays/Year/Month
        JPanel step2 = new JPanel(new GridBagLayout());
        gbc.gridy = 0;
        step2.add(new JLabel("Holidays (comma-separated days, e.g., 15,26):"), gbc);
        gbc.gridy++;
        holidaysField = new JTextField("15,26");
        step2.add(holidaysField, gbc);

        gbc.gridy++;
        step2.add(new JLabel("Year (auto-detected if blank):"), gbc);
        gbc.gridy++;
        yearField = new JTextField();
        step2.add(yearField, gbc);

        gbc.gridy++;
        step2.add(new JLabel("Month (1-12, auto-detected if blank):"), gbc);
        gbc.gridy++;
        monthField = new JTextField();
        step2.add(monthField, gbc);

        gbc.gridy++;
        JPanel buttonPanel = new JPanel(new FlowLayout());
        JButton backBtn = new JButton("Back");
        backBtn.addActionListener(e -> wizardLayout.previous(wizardPanel));
        processBtn = new JButton("Process");
        processBtn.addActionListener(this::processData);
        exportCsvBtn = new JButton("Export Report to CSV");
        exportCsvBtn.setEnabled(false);
        exportCsvBtn.addActionListener(this::exportCsv);
        buttonPanel.add(backBtn);
        buttonPanel.add(processBtn);
        buttonPanel.add(exportCsvBtn);
        step2.add(buttonPanel, gbc);

        wizardPanel.add(step2, "STEP2");

        return wizardPanel;
    }

    private JPanel createDataViewerPanel() {
        JPanel viewerPanel = new JPanel(new BorderLayout());
        JPanel queryPanel = new JPanel(new FlowLayout());
        queryPanel.add(new JLabel("Employee Code (blank for all):"));
        empCodeField = new JTextField(10);
        queryPanel.add(empCodeField);

        queryPanel.add(new JLabel("From Day:"));
        fromDayField = new JTextField(3);
        queryPanel.add(fromDayField);

        queryPanel.add(new JLabel("To Day:"));
        toDayField = new JTextField(3);
        queryPanel.add(toDayField);

        queryBtn = new JButton("Query Data");
        queryBtn.addActionListener(this::queryData);
        queryBtn.setEnabled(false);
        queryPanel.add(queryBtn);

        viewerPanel.add(queryPanel, BorderLayout.NORTH);

        dataModel = new DefaultTableModel();
        dataTable = new JTable(dataModel);
        dataTable.setToolTipText("Attendance data table");
        viewerPanel.add(new JScrollPane(dataTable), BorderLayout.CENTER);

        return viewerPanel;
    }

    private JPanel createReportViewerPanel() {
        JPanel reportViewerPanel = new JPanel(new BorderLayout());
        reportModel = new DefaultTableModel();
        reportTable = new JTable(reportModel);
        reportTable.setToolTipText("Employee metrics report");
        reportViewerPanel.add(new JScrollPane(reportTable), BorderLayout.CENTER);
        return reportViewerPanel;
    }

    private void selectFile(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        int res = chooser.showOpenDialog(this);
        if (res == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();
            if (e.getSource() == selectInputBtn) inputFileField.setText(file.getAbsolutePath());
            else if (e.getSource() == selectMergedBtn) mergedFileField.setText(file.getAbsolutePath());
            else if (e.getSource() == selectReportBtn) reportFileField.setText(file.getAbsolutePath());
        }
    }

    private void processData(ActionEvent e) {
        try {
            String inputFile = inputFileField.getText().trim();
            String mergedFile = mergedFileField.getText().trim();
            String reportFile = reportFileField.getText().trim();
            if (inputFile.isEmpty() || mergedFile.isEmpty() || reportFile.isEmpty()) {
                JOptionPane.showMessageDialog(this, "Please select all files.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            List<Integer> holidays = Arrays.stream(holidaysField.getText().split(","))
                    .filter(s -> !s.trim().isEmpty())
                    .map(Integer::parseInt)
                    .collect(Collectors.toList());

            mergeResult = AttendanceMerger.merge(inputFile, mergedFile, holidays);

            if (!yearField.getText().isEmpty()) mergeResult.year = Integer.parseInt(yearField.getText());
            if (!monthField.getText().isEmpty()) mergeResult.month = Integer.parseInt(monthField.getText());

            reportGenerator.generate(mergedFile, reportFile, mergeResult.allEmployees, mergeResult.monthDays, holidays, mergeResult.year, mergeResult.month);

            loadReportIntoTable(reportFile);

            exportCsvBtn.setEnabled(true);
            queryBtn.setEnabled(true);

            queryViewer = new AttendanceQueryViewer(mergedFile, mergeResult.allEmployees);

            JOptionPane.showMessageDialog(this, "Processing complete! Check Data Viewer and Report Viewer.");
            showPanel("DATA"); // Auto-switch to Data Viewer
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(this, "Error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void exportCsv(ActionEvent e) {
        if (mergeResult == null) return;
        JFileChooser chooser = new JFileChooser();
        chooser.setSelectedFile(new File("report.csv"));
        int res = chooser.showSaveDialog(this);
        if (res == JFileChooser.APPROVE_OPTION) {
            try {
                reportGenerator.exportToCsv(chooser.getSelectedFile().getAbsolutePath(), mergeResult.allEmployees,
                        mergeResult.monthDays, parseHolidays(), mergeResult.year, mergeResult.month);
                JOptionPane.showMessageDialog(this, "CSV exported successfully!");
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Error exporting CSV: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void queryData(ActionEvent e) {
        if (mergeResult == null) {
            JOptionPane.showMessageDialog(this, "Process data first!");
            return;
        }
        dataModel.setRowCount(0);
        dataModel.setColumnIdentifiers(new Object[0]);

        String empCode = empCodeField.getText().trim();
        int from = fromDayField.getText().isEmpty() ? 1 : Integer.parseInt(fromDayField.getText());
        int to = toDayField.getText().isEmpty() ? mergeResult.monthDays : Integer.parseInt(toDayField.getText());

        if (!empCode.isEmpty()) {
            queryViewer.showByEmployee(empCode);
            return;
        } else if (from == to) {
            queryViewer.showByDate(from);
            return;
        } else {
            Object[] columns = {"Employee Code", "Employee Name", "Day", "InTime", "OutTime", "Duration", "Status"};
            dataModel.setColumnIdentifiers(columns);

            for (AttendanceMerger.EmployeeData emp : mergeResult.allEmployees) {
                for (int d = from; d <= to; d++) {
                    String in = AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d - 1);
                    String out = AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), d - 1);
                    String dur = AttendanceUtils.safeGet(emp.dailyData.get("Duration"), d - 1);
                    String stat = AttendanceUtils.safeGet(emp.dailyData.get("Status"), d - 1);

                    dataModel.addRow(new Object[]{emp.empId, emp.empName, d, in, out, dur, stat});
                }
            }
        }
    }

    private void loadReportIntoTable(String reportFile) throws Exception {
        FileInputStream fis = new FileInputStream(reportFile);
        Workbook wb = new XSSFWorkbook(fis);
        Sheet sheet = wb.getSheet("Employee_Report");

        reportModel.setRowCount(0);
        reportModel.setColumnIdentifiers(new Object[0]);

        Row header = sheet.getRow(0);
        Object[] columns = new Object[header.getLastCellNum()];
        for (int i = 0; i < columns.length; i++) {
            columns[i] = header.getCell(i).getStringCellValue();
        }
        reportModel.setColumnIdentifiers(columns);

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            Object[] data = new Object[row.getLastCellNum()];
            for (int c = 0; c < data.length; c++) {
                Cell cell = row.getCell(c);
                if (cell.getCellType() == CellType.NUMERIC) data[c] = cell.getNumericCellValue();
                else data[c] = cell.getStringCellValue();
            }
            reportModel.addRow(data);
        }

        wb.close();
        fis.close();
    }

    private List<Integer> parseHolidays() {
        return Arrays.stream(holidaysField.getText().split(","))
                .filter(s -> !s.trim().isEmpty())
                .map(Integer::parseInt)
                .collect(Collectors.toList());
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new AttendanceAppGUI().setVisible(true));
    }
}