package org.bioparse.cleaning.ui;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.table.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.util.*;
import java.util.List;
import java.util.regex.Pattern;

public class ReportViewerPanel extends JPanel {
    
    private static final long serialVersionUID = 1L;
    
    private JTable reportTable;
    private DefaultTableModel reportModel;
    private JTextField globalSearchField;
    private JComboBox<String> deptFilterCombo;
    private TableRowSorter<DefaultTableModel> reportSorter;
    private JLabel statsLabel;
    
    private JPanel reportFilterPanel;
    
    public ReportViewerPanel() {
        setLayout(new BorderLayout(0, 10));
        setBackground(new Color(245, 247, 250));
        setBorder(new EmptyBorder(10, 10, 10, 10));
        
        setupHeader();
        setupFilterPanel();
        setupTable();
        
        // Initially show empty state
        showEmptyState();
    }
    
    private void setupHeader() {
        JPanel header = new JPanel(new BorderLayout());
        header.setBackground(Color.WHITE);
        header.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createMatteBorder(0, 0, 1, 0, new Color(225, 228, 232)),
            BorderFactory.createEmptyBorder(10, 15, 10, 15)
        ));
        
        JLabel title = new JLabel("📋 Employee Report - Monthly Summary");
        title.setFont(new Font("Segoe UI Semibold", Font.PLAIN, 18));
        title.setForeground(new Color(52, 58, 64));
        
        statsLabel = new JLabel("No data loaded");
        statsLabel.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        statsLabel.setForeground(new Color(108, 117, 125));
        
        header.add(title, BorderLayout.WEST);
        header.add(statsLabel, BorderLayout.EAST);
        add(header, BorderLayout.NORTH);
    }
    
    private void setupFilterPanel() {
        reportFilterPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
        reportFilterPanel.setBackground(new Color(255, 255, 255, 220));
        reportFilterPanel.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(233, 236, 239), 1),
            BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));
        
        // Search field
        JLabel searchLabel = new JLabel("🔍 Search:");
        searchLabel.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        globalSearchField = new JTextField(20);
        globalSearchField.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        globalSearchField.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(206, 212, 218)),
            BorderFactory.createEmptyBorder(5, 8, 5, 8)
        ));
        
        // Department filter
        JLabel deptLabel = new JLabel("Department:");
        deptLabel.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        deptFilterCombo = new JComboBox<>(new String[]{"All", "IT", "HR", "Finance", "Operations", "Admin", "Sales", "Marketing"});
        deptFilterCombo.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        deptFilterCombo.setPreferredSize(new Dimension(120, 30));
        
        // Export button
        JButton exportBtn = new JButton("📥 Export CSV");
        exportBtn.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        exportBtn.setBackground(new Color(40, 167, 69));
        exportBtn.setForeground(Color.WHITE);
        exportBtn.setBorder(BorderFactory.createEmptyBorder(8, 15, 8, 15));
        exportBtn.setFocusPainted(false);
        exportBtn.setCursor(new Cursor(Cursor.HAND_CURSOR));
        exportBtn.addActionListener(e -> exportReport());
        
        // Clear filters button
        JButton clearBtn = new JButton("🗑️ Clear Filters");
        clearBtn.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        clearBtn.setBackground(new Color(108, 117, 125));
        clearBtn.setForeground(Color.WHITE);
        clearBtn.setBorder(BorderFactory.createEmptyBorder(8, 15, 8, 15));
        clearBtn.setFocusPainted(false);
        clearBtn.setCursor(new Cursor(Cursor.HAND_CURSOR));
        clearBtn.addActionListener(e -> clearFilters());
        
        reportFilterPanel.add(searchLabel);
        reportFilterPanel.add(globalSearchField);
        reportFilterPanel.add(Box.createHorizontalStrut(20));
        reportFilterPanel.add(deptLabel);
        reportFilterPanel.add(deptFilterCombo);
        reportFilterPanel.add(Box.createHorizontalStrut(20));
        reportFilterPanel.add(exportBtn);
        reportFilterPanel.add(clearBtn);
        
        reportFilterPanel.setVisible(false);
        add(reportFilterPanel, BorderLayout.CENTER);
    }
    
    private void setupTable() {
        reportModel = new DefaultTableModel() {
            private static final long serialVersionUID = 1L;
            
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // Make table read-only
            }
            
            @Override
            public Class<?> getColumnClass(int columnIndex) {
                // All columns are String type to avoid DoubleRenderer issues
                return String.class;
            }
        };
        
        reportTable = new JTable(reportModel);
        configureTableAppearance();
        
        JScrollPane scrollPane = new JScrollPane(reportTable);
        scrollPane.setBorder(BorderFactory.createEmptyBorder());
        scrollPane.getViewport().setBackground(Color.WHITE);
        
        add(scrollPane, BorderLayout.SOUTH);
    }
    
    private void configureTableAppearance() {
        reportTable.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        reportTable.setRowHeight(32);
        reportTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        reportTable.setGridColor(new Color(233, 236, 239));
        reportTable.setShowGrid(true);
        reportTable.setIntercellSpacing(new Dimension(0, 1));
        reportTable.setFillsViewportHeight(true);
        
        // Header styling
        JTableHeader header = reportTable.getTableHeader();
        header.setFont(new Font("Segoe UI Semibold", Font.PLAIN, 12));
        header.setBackground(new Color(52, 58, 64));
        header.setForeground(Color.WHITE);
        header.setReorderingAllowed(false);
        
        // Custom renderer for all cells
        reportTable.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            private static final long serialVersionUID = 1L;
            
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, 
                    boolean isSelected, boolean hasFocus, int row, int column) {
                
                String text = "";
                if (value != null) {
                    text = value.toString();
                }
                
                Component c = super.getTableCellRendererComponent(table, text, 
                        isSelected, hasFocus, row, column);
                
                if (text.isEmpty() || "-".equals(text)) {
                    setText("-");
                    setForeground(new Color(173, 181, 189));
                } else {
                    setText(text);
                    
                    // Highlight problematic values (absences, lates, etc.)
                    if (column >= 7 && column <= 9) { // Lates, Absent, Punch Missed columns
                        try {
                            int val = Integer.parseInt(text);
                            if (val > 0) {
                                setForeground(isSelected ? Color.WHITE : new Color(220, 53, 69)); // Red for issues
                            } else {
                                setForeground(isSelected ? Color.WHITE : new Color(33, 37, 41));
                            }
                        } catch (NumberFormatException e) {
                            setForeground(isSelected ? Color.WHITE : new Color(33, 37, 41));
                        }
                    } else if (column == 12) { // OT Hours column
                        try {
                            double otHours = Double.parseDouble(text);
                            if (otHours > 0) {
                                setForeground(isSelected ? Color.WHITE : new Color(40, 167, 69)); // Green for OT
                            } else {
                                setForeground(isSelected ? Color.WHITE : new Color(33, 37, 41));
                            }
                        } catch (NumberFormatException e) {
                            setForeground(isSelected ? Color.WHITE : new Color(33, 37, 41));
                        }
                    } else {
                        setForeground(isSelected ? Color.WHITE : new Color(33, 37, 41));
                    }
                }
                
                setHorizontalAlignment(getAlignmentForColumn(column));
                
                // Alternating row colors
                if (!isSelected) {
                    c.setBackground(row % 2 == 0 ? Color.WHITE : new Color(248, 249, 250));
                } else {
                    c.setBackground(new Color(0, 123, 255));
                }
                
                setBorder(BorderFactory.createEmptyBorder(0, 8, 0, 8));
                return c;
            }
            
            private int getAlignmentForColumn(int column) {
                if (column == 0) { // Employee Code
                    return SwingConstants.LEFT;
                } else if (column >= 1 && column <= 11) { // Numeric columns
                    return SwingConstants.CENTER;
                } else if (column == 12) { // OT Hours (decimal)
                    return SwingConstants.RIGHT;
                }
                return SwingConstants.LEFT;
            }
        });
        
        // Enable column sorting with custom comparator
        reportTable.setAutoCreateRowSorter(true);
        
        // Set custom comparator for numeric columns
        setupColumnSorting();
    }
    
    private void setupColumnSorting() {
        if (reportTable.getRowSorter() != null) {
            TableRowSorter<DefaultTableModel> sorter = (TableRowSorter<DefaultTableModel>) reportTable.getRowSorter();
            
            // Add numeric comparators for columns that should be sorted as numbers
            for (int col = 1; col < reportModel.getColumnCount(); col++) {
                sorter.setComparator(col, new Comparator<String>() {
                    @Override
                    public int compare(String s1, String s2) {
                        try {
                            // Try to parse as double first
                            double d1 = Double.parseDouble(s1);
                            double d2 = Double.parseDouble(s2);
                            return Double.compare(d1, d2);
                        } catch (NumberFormatException e) {
                            // If not numeric, compare as strings
                            return s1.compareTo(s2);
                        }
                    }
                });
            }
        }
    }
    
    private void showEmptyState() {
        reportModel.setRowCount(0);
        reportModel.setColumnCount(0);
        
        JPanel emptyPanel = new JPanel(new BorderLayout());
        emptyPanel.setBackground(Color.WHITE);
        emptyPanel.setBorder(BorderFactory.createEmptyBorder(100, 20, 100, 20));
        
        JLabel emptyLabel = new JLabel(
            "<html><div style='text-align: center;'>" +
            "<div style='font-size: 48px; color: #adb5bd; margin-bottom: 20px;'>📄</div>" +
            "<div style='font-size: 18px; color: #495057; margin-bottom: 10px;'>No Report Data</div>" +
            "<div style='font-size: 14px; color: #6c757d;'>Process attendance data first to view monthly employee summaries</div>" +
            "</div></html>", SwingConstants.CENTER);
        
        // Remove existing components and add empty state
        removeAll();
        setupHeader();
        add(emptyPanel, BorderLayout.CENTER);
        emptyPanel.add(emptyLabel, BorderLayout.CENTER);
        
        revalidate();
        repaint();
    }
    
    /**
     * Load report data directly (avoiding Excel reading issues)
     */
    public void loadReportData(String[][] data, String[] columns) {
        // Validate data
        if (data == null || columns == null || data.length == 0) {
            showEmptyState();
            return;
        }
        
        // Clean and validate data before loading
        String[][] cleanedData = cleanData(data, columns);
        
        // Remove existing components and show data
        removeAll();
        setupHeader();
        setupFilterPanel();
        
        // Load data into model
        reportModel.setDataVector(cleanedData, columns);
        
        // Setup table and add to panel
        configureTableAppearance();
        JScrollPane scrollPane = new JScrollPane(reportTable);
        scrollPane.setBorder(BorderFactory.createEmptyBorder());
        scrollPane.getViewport().setBackground(Color.WHITE);
        
        add(reportFilterPanel, BorderLayout.CENTER);
        add(scrollPane, BorderLayout.SOUTH);
        
        // Setup column sorting
        setupColumnSorting();
        
        // Update stats
        updateStats();
        
        // Setup filtering
        setupFilters();
        
        // Show filter panel
        reportFilterPanel.setVisible(true);
        
        revalidate();
        repaint();
    }
    
    /**
     * Clean and validate data to prevent rendering errors
     */
    private String[][] cleanData(String[][] data, String[] columns) {
        String[][] cleaned = new String[data.length][columns.length];
        
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < columns.length; j++) {
                String value = (j < data[i].length) ? data[i][j] : "";
                
                // Clean the value
                if (value == null || value.trim().isEmpty()) {
                    cleaned[i][j] = "0"; // Default to 0 for numeric columns
                } else {
                    value = value.trim();
                    
                    // Ensure OT Hours column has proper decimal format
                    if (j == columns.length - 1 && columns[j].toLowerCase().contains("ot hours")) {
                        try {
                            double d = Double.parseDouble(value);
                            cleaned[i][j] = String.format("%.2f", d);
                        } catch (NumberFormatException e) {
                            cleaned[i][j] = "0.00"; // Default to 0.00
                        }
                    } else {
                        cleaned[i][j] = value;
                    }
                }
            }
        }
        
        return cleaned;
    }
    
    private void updateStats() {
        int totalEmployees = reportModel.getRowCount();
        
        if (totalEmployees == 0) {
            statsLabel.setText("No data loaded");
            return;
        }
        
        // Calculate some statistics
        int totalAbsent = 0;
        int totalLates = 0;
        double totalOTHours = 0;
        
        for (int row = 0; row < reportModel.getRowCount(); row++) {
            try {
                Object lateValue = reportModel.getValueAt(row, 7);
                if (lateValue != null) {
                    totalLates += Integer.parseInt(lateValue.toString());
                }
            } catch (Exception e) {
                // Ignore parsing errors
            }
            
            try {
                Object absentValue = reportModel.getValueAt(row, 8);
                if (absentValue != null) {
                    totalAbsent += Integer.parseInt(absentValue.toString());
                }
            } catch (Exception e) {
                // Ignore parsing errors
            }
            
            try {
                Object otValue = reportModel.getValueAt(row, 12);
                if (otValue != null) {
                    totalOTHours += Double.parseDouble(otValue.toString());
                }
            } catch (Exception e) {
                // Ignore parsing errors
            }
        }
        
        String stats = String.format("📊 %d employees", totalEmployees);
        if (totalAbsent > 0) stats += String.format(" | %d absences", totalAbsent);
        if (totalLates > 0) stats += String.format(" | %d lates", totalLates);
        if (totalOTHours > 0) stats += String.format(" | %.1f OT hours", totalOTHours);
        
        statsLabel.setText(stats);
    }
    
    private void setupFilters() {
        reportSorter = new TableRowSorter<>(reportModel);
        reportTable.setRowSorter(reportSorter);
        
        // Setup numeric comparators for filtering
        for (int col = 1; col < reportModel.getColumnCount(); col++) {
            reportSorter.setComparator(col, new Comparator<String>() {
                @Override
                public int compare(String s1, String s2) {
                    try {
                        double d1 = Double.parseDouble(s1);
                        double d2 = Double.parseDouble(s2);
                        return Double.compare(d1, d2);
                    } catch (NumberFormatException e) {
                        return s1.compareTo(s2);
                    }
                }
            });
        }
        
        // Clear existing filters
        globalSearchField.setText("");
        deptFilterCombo.setSelectedIndex(0);
        
        // Setup search filter
        globalSearchField.getDocument().addDocumentListener(new DocumentListener() {
            @Override public void insertUpdate(DocumentEvent e) { applyFilters(); }
            @Override public void removeUpdate(DocumentEvent e) { applyFilters(); }
            @Override public void changedUpdate(DocumentEvent e) { applyFilters(); }
        });
        
        // Add filter listeners
        deptFilterCombo.addActionListener(e -> applyFilters());
    }
    
    private void applyFilters() {
        if (reportSorter == null) return;
        
        List<RowFilter<Object, Object>> filters = new ArrayList<>();
        
        // Text search filter
        String searchText = globalSearchField.getText().trim();
        if (!searchText.isEmpty()) {
            String regex = "(?i).*" + Pattern.quote(searchText) + ".*";
            filters.add(RowFilter.regexFilter(regex));
        }
        
        // Department filter
        String selectedDept = (String) deptFilterCombo.getSelectedItem();
        if (selectedDept != null && !"All".equals(selectedDept)) {
            // Find employee code column (first column)
            filters.add(RowFilter.regexFilter("(?i).*" + Pattern.quote(selectedDept) + ".*", 0));
        }
        
        // Apply combined filter
        if (filters.isEmpty()) {
            reportSorter.setRowFilter(null);
        } else {
            reportSorter.setRowFilter(RowFilter.andFilter(filters));
        }
    }
    
    private void clearFilters() {
        globalSearchField.setText("");
        deptFilterCombo.setSelectedIndex(0);
        if (reportSorter != null) {
            reportSorter.setRowFilter(null);
        }
    }
    
    private void exportReport() {
        if (reportModel.getRowCount() == 0) {
            JOptionPane.showMessageDialog(this,
                "No data to export. Please generate a report first.",
                "No Data", JOptionPane.WARNING_MESSAGE);
            return;
        }
        
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Export Monthly Report");
        fileChooser.setSelectedFile(new File("Monthly_Attendance_Report.csv"));
        
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            try {
                exportToCsv(file);
                JOptionPane.showMessageDialog(this,
                    "Report exported successfully to:\n" + file.getAbsolutePath(),
                    "Export Complete", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(this,
                    "Error exporting report: " + e.getMessage(),
                    "Export Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }
    
    private void exportToCsv(File file) throws Exception {
        try (java.io.PrintWriter writer = new java.io.PrintWriter(file)) {
            // Write headers
            for (int i = 0; i < reportModel.getColumnCount(); i++) {
                writer.print(reportModel.getColumnName(i));
                if (i < reportModel.getColumnCount() - 1) writer.print(",");
            }
            writer.println();
            
            // Write data
            for (int row = 0; row < reportModel.getRowCount(); row++) {
                for (int col = 0; col < reportModel.getColumnCount(); col++) {
                    Object value = reportModel.getValueAt(row, col);
                    String cellValue = value != null ? value.toString() : "";
                    // Escape quotes and commas
                    if (cellValue.contains(",") || cellValue.contains("\"")) {
                        cellValue = "\"" + cellValue.replace("\"", "\"\"") + "\"";
                    }
                    writer.print(cellValue);
                    if (col < reportModel.getColumnCount() - 1) writer.print(",");
                }
                writer.println();
            }
        }
    }
    
    // Method to clear all data
    public void clearData() {
        showEmptyState();
    }
    
    // Method to check if data is loaded
    public boolean hasData() {
        return reportModel.getRowCount() > 0;
    }
    
    // Getter for table
    public JTable getReportTable() {
        return reportTable;
    }
    
    // Getter for model
    public DefaultTableModel getReportModel() {
        return reportModel;
    }
}