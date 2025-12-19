package org.bioparse.cleaning.ui;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.util.List;
import java.util.Map;

import org.bioparse.cleaning.AttendanceMerger;
import org.bioparse.cleaning.AttendanceReportGenerator;
import org.bioparse.cleaning.Constants;

public class DashboardPanel extends JPanel {
    private static final long serialVersionUID = 1L;
    private JLabel totalEmployeesLabel, presentDaysLabel, totalAbsentLabel, totalLatesLabel;
    private JLabel totalOTHoursLabel, totalHalfDaysLabel, totalWorkingDaysLabel, monthLabel;
    private JPanel statsGridPanel;
    private AttendanceMerger.MergeResult mergeResult;
    private int selectedYear, selectedMonth;
    
    public DashboardPanel() {
        setLayout(new BorderLayout());
        setBackground(Constants.BACKGROUND_COLOR);
        setBorder(new EmptyBorder(20, 20, 20, 20));
        
        setupHeader();
        setupContent();
    }
    
    private void setupHeader() {
        JPanel headerPanel = new JPanel(new BorderLayout());
        headerPanel.setBackground(Constants.BACKGROUND_COLOR);
        headerPanel.setBorder(new EmptyBorder(0, 0, 30, 0));
        
        JLabel headerLabel = new JLabel("Attendance Analytics Dashboard");
        headerLabel.setFont(new Font("Segoe UI", Font.BOLD, 28));
        headerLabel.setForeground(Constants.TEXT_PRIMARY);
        
        monthLabel = new JLabel("No data processed yet");
        monthLabel.setFont(Constants.FONT_SUBHEADER);
        monthLabel.setForeground(Constants.TEXT_SECONDARY);
        monthLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        
        headerPanel.add(headerLabel, BorderLayout.WEST);
        headerPanel.add(monthLabel, BorderLayout.EAST);
        
        add(headerPanel, BorderLayout.NORTH);
    }
    
    private void setupContent() {
        JPanel contentPanel = new JPanel(new GridBagLayout());
        contentPanel.setBackground(Constants.BACKGROUND_COLOR);
        GridBagConstraints gbc = new GridBagConstraints();
        
        // Left side - Stats cards in grid
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0.7;
        gbc.weighty = 1.0;
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(0, 0, 0, 20);
        
        statsGridPanel = createStatsGrid();
        contentPanel.add(statsGridPanel, gbc);
        
        // Right side - Summary and quick actions
        gbc.gridx = 1;
        gbc.gridy = 0;
        gbc.weightx = 0.3;
        gbc.insets = new Insets(0, 0, 0, 0);
        
        JPanel summaryPanel = createSummaryPanel();
        contentPanel.add(summaryPanel, gbc);
        
        add(contentPanel, BorderLayout.CENTER);
    }
    
    private JPanel createStatsGrid() {
        JPanel panel = new JPanel(new GridLayout(2, 4, 15, 15));
        panel.setBackground(Constants.BACKGROUND_COLOR);
        
        // Row 1
        totalEmployeesLabel = createStatCard("Total Employees", "0", Constants.PRIMARY_COLOR, "👥");
        presentDaysLabel = createStatCard("Present Days", "0", Constants.SUCCESS_COLOR, "✅");
        totalAbsentLabel = createStatCard("Total Absent", "0", Constants.WARNING_COLOR, "❌");
        totalLatesLabel = createStatCard("Late Arrivals", "0", Constants.DANGER_COLOR, "⏰");
        
        // Row 2
        totalOTHoursLabel = createStatCard("OT Hours", "0", new Color(155, 89, 182), "⏱️");
        totalHalfDaysLabel = createStatCard("Half Days", "0", new Color(243, 156, 18), "½");
        totalWorkingDaysLabel = createStatCard("Working Days", "0", new Color(52, 152, 219), "📅");
        
        // Empty card for balance
        JLabel emptyCard = createStatCard("Month", "No Data", new Color(108, 117, 125), "📊");
        
        panel.add(totalEmployeesLabel);
        panel.add(presentDaysLabel);
        panel.add(totalAbsentLabel);
        panel.add(totalLatesLabel);
        panel.add(totalOTHoursLabel);
        panel.add(totalHalfDaysLabel);
        panel.add(totalWorkingDaysLabel);
        panel.add(emptyCard);
        
        return panel;
    }
    
    private JLabel createStatCard(String title, String value, Color color, String emoji) {
        String html = "<html><div style='background: white; border-left: 4px solid " + 
                     String.format("#%02x%02x%02x", color.getRed(), color.getGreen(), color.getBlue()) + 
                     "; border-radius: 8px; padding: 20px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>" +
                     "<div style='display: flex; align-items: center; margin-bottom: 10px;'>" +
                     "<span style='font-size: 24px; margin-right: 10px;'>" + emoji + "</span>" +
                     "<span style='font-size: 14px; color: #666;'>" + title + "</span>" +
                     "</div>" +
                     "<div style='font-size: 32px; font-weight: bold; color: " +
                     String.format("#%02x%02x%02x", color.getRed(), color.getGreen(), color.getBlue()) + 
                     "; margin-top: 5px;'>" + value + "</div>" +
                     "</div></html>";
        
        JLabel card = new JLabel(html);
        card.setOpaque(false);
        card.setToolTipText(title);
        return card;
    }
    
    private JPanel createSummaryPanel() {
        JPanel panel = new JPanel(new BorderLayout(0, 20));
        panel.setBackground(Constants.CARD_BACKGROUND);
        panel.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(Constants.BORDER_COLOR, 1),
            new EmptyBorder(20, 20, 20, 20)
        ));
        
        // Summary header
        JLabel summaryHeader = new JLabel("📋 Month Summary");
        summaryHeader.setFont(Constants.FONT_HEADER);
        summaryHeader.setForeground(Constants.TEXT_PRIMARY);
        summaryHeader.setBorder(new EmptyBorder(0, 0, 10, 0));
        
        // Summary content
        JTextArea summaryText = new JTextArea();
        summaryText.setText("Process attendance data to see monthly summary.\n\n" +
                           "1. Go to Configuration panel\n" +
                           "2. Select input file\n" +
                           "3. Choose month & year\n" +
                           "4. Click 'Merge & Process Data'");
        summaryText.setFont(Constants.FONT_BODY);
        summaryText.setForeground(Constants.TEXT_SECONDARY);
        summaryText.setWrapStyleWord(true);
        summaryText.setLineWrap(true);
        summaryText.setEditable(false);
        summaryText.setBackground(Constants.CARD_BACKGROUND);
        summaryText.setBorder(new EmptyBorder(10, 10, 10, 10));
        
        panel.add(summaryHeader, BorderLayout.NORTH);
        panel.add(summaryText, BorderLayout.CENTER);
        
        return panel;
    }
    
    public void updateStats(AttendanceMerger.MergeResult mergeResult) {
        this.mergeResult = mergeResult;
        if (mergeResult != null) {
            calculateAndDisplayStats();
        }
    }
    
    public void setSelectedMonthYear(int year, int month) {
        this.selectedYear = year;
        this.selectedMonth = month;
        updateMonthLabel();
    }
    
    private void updateMonthLabel() {
        if (selectedYear > 0 && selectedMonth > 0) {
            String monthName = getMonthName(selectedMonth);
            monthLabel.setText(monthName + " " + selectedYear);
        }
    }
    
private void calculateAndDisplayStats() {
    if (mergeResult == null || mergeResult.allEmployees.isEmpty()) {
        resetStats();
        return;
    }
    
    int totalEmployees = mergeResult.allEmployees.size();
    int totalWorkingDays = mergeResult.monthDays;
    
    // Use AttendanceReportGenerator to calculate accurate metrics
    AttendanceReportGenerator reportGenerator = new AttendanceReportGenerator();
    
    int totalFullDays = 0;
    int totalAbsent = 0;
    int totalLates = 0;
    int totalHalfDays = 0;
    double totalOTHours = 0;
    int totalPunchMissed = 0;
    
    // Convert List<Object> to List<AttendanceMerger.EmployeeData>
    List<AttendanceMerger.EmployeeData> employees = new java.util.ArrayList<>();
    for (Object empObj : mergeResult.allEmployees) {
        if (empObj instanceof AttendanceMerger.EmployeeData) {
            employees.add((AttendanceMerger.EmployeeData) empObj);
        }
    }
    
    // Process each employee
    for (AttendanceMerger.EmployeeData emp : employees) {
        // Calculate metrics using the report generator
        AttendanceReportGenerator.Metrics metrics = reportGenerator.computeMetrics(
            emp, 
            totalWorkingDays, 
            java.util.Collections.emptyList(), // Pass empty holidays for dashboard
            mergeResult.year, 
            mergeResult.month
        );
        
        // Accumulate statistics
        totalFullDays += metrics.totalFullDays;
        totalAbsent += metrics.totalAbsent;
        totalLates += metrics.totalLates;
        totalHalfDays += metrics.halfDays;
        totalPunchMissed += metrics.totalPunchMissed;
        
        // Calculate OT hours (same logic as report generator)
        // Working-day OT (rounded to nearest 30 min)
        int blocks = (metrics.workingDayOTMinutes + 29) / 30;
        double workingOTHours = blocks * 0.5;
        
        // Weekend/Holiday OT
        double weekendFullOTHours = metrics.weekendFullOTMinutes / 60.0;
        double weekendHalfOTHours = metrics.weekendHalfOTMinutes / 60.0;
        
        totalOTHours += workingOTHours + weekendFullOTHours + weekendHalfOTHours;
    }
    
    // Debug output to console
    System.out.println("\n=== DASHBOARD STATISTICS ===");
    System.out.println("Total Employees: " + totalEmployees);
    System.out.println("Month: " + mergeResult.month + "/" + mergeResult.year);
    System.out.println("Working Days in Month: " + totalWorkingDays);
    System.out.println("Total Person-Days: " + (totalEmployees * totalWorkingDays));
    System.out.println("---");
    System.out.println("Present (Full) Days: " + totalFullDays);
    System.out.println("Absent Days: " + totalAbsent);
    System.out.println("Half Days: " + totalHalfDays);
    System.out.println("Late Arrivals: " + totalLates);
    System.out.println("Punch Missed: " + totalPunchMissed);
    System.out.println("OT Hours: " + String.format("%.1f", totalOTHours));
    System.out.println("=== END STATISTICS ===\n");
    
    // Calculate accounted days for validation
    int accountedDays = totalFullDays + totalAbsent + totalHalfDays;
    int totalPersonDays = totalEmployees * totalWorkingDays;
    
    if (accountedDays < totalPersonDays) {
        System.out.println("Note: " + (totalPersonDays - accountedDays) + " days not accounted for " +
                          "(may be weekends/holidays or other statuses)");
    }
    
    // Update UI
    updateStatCard(totalEmployeesLabel, String.valueOf(totalEmployees));
    updateStatCard(presentDaysLabel, String.valueOf(totalFullDays));
    updateStatCard(totalAbsentLabel, String.valueOf(totalAbsent));
    updateStatCard(totalLatesLabel, String.valueOf(totalLates));
    updateStatCard(totalHalfDaysLabel, String.valueOf(totalHalfDays));
    updateStatCard(totalWorkingDaysLabel, String.valueOf(totalWorkingDays));
    updateStatCard(totalOTHoursLabel, String.format("%.1f", totalOTHours));
    
    // Update month label
    if (mergeResult.year > 0 && mergeResult.month > 0) {
        String monthName = getMonthName(mergeResult.month);
        monthLabel.setText(monthName + " " + mergeResult.year + " • " + totalEmployees + " employees");
    }
}  
    private boolean isLate(String inTime) {
        try {
            String[] parts = inTime.split(":");
            if (parts.length >= 2) {
                int hour = Integer.parseInt(parts[0]);
                int minute = Integer.parseInt(parts[1]);
                return (hour > 9) || (hour == 9 && minute > 45);
            }
        } catch (Exception e) {
            // If can't parse, assume not late
        }
        return false;
    }
    
    private double parseOTHours(String otString) {
        try {
            if (otString.contains(":")) {
                String[] parts = otString.split(":");
                int hours = Integer.parseInt(parts[0]);
                int minutes = parts.length > 1 ? Integer.parseInt(parts[1]) : 0;
                return hours + (minutes / 60.0);
            } else {
                return Double.parseDouble(otString);
            }
        } catch (Exception e) {
            return 0.0;
        }
    }
    
    private void resetStats() {
        updateStatCard(totalEmployeesLabel, "0");
        updateStatCard(presentDaysLabel, "0");
        updateStatCard(totalAbsentLabel, "0");
        updateStatCard(totalLatesLabel, "0");
        updateStatCard(totalOTHoursLabel, "0");
        updateStatCard(totalHalfDaysLabel, "0");
        updateStatCard(totalWorkingDaysLabel, "0");
        monthLabel.setText("No data processed yet");
    }
    
    private void updateStatCard(JLabel card, String value) {
        String currentHtml = card.getText();
        String emoji = extractEmoji(currentHtml);
        Color color = extractColor(currentHtml);
        String title = extractTitle(currentHtml);
        
        String newHtml = "<html><div style='background: white; border-left: 4px solid " + 
                        String.format("#%02x%02x%02x", color.getRed(), color.getGreen(), color.getBlue()) + 
                        "; border-radius: 8px; padding: 20px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>" +
                        "<div style='display: flex; align-items: center; margin-bottom: 10px;'>" +
                        "<span style='font-size: 24px; margin-right: 10px;'>" + emoji + "</span>" +
                        "<span style='font-size: 14px; color: #666;'>" + title + "</span>" +
                        "</div>" +
                        "<div style='font-size: 32px; font-weight: bold; color: " +
                        String.format("#%02x%02x%02x", color.getRed(), color.getGreen(), color.getBlue()) + 
                        "; margin-top: 5px;'>" + value + "</div>" +
                        "</div></html>";
        
        card.setText(newHtml);
        card.setToolTipText(title);
        card.revalidate();
        card.repaint();
    }
    
    private String extractEmoji(String html) {
        int start = html.indexOf("font-size: 24px; margin-right: 10px;'>");
        if (start != -1) {
            start += 38;
            int end = html.indexOf("</span>", start);
            if (end > start) {
                return html.substring(start, end).trim();
            }
        }
        return "📊";
    }
    
    private String extractTitle(String html) {
        int start = html.indexOf("font-size: 14px; color: #666;'>");
        if (start != -1) {
            start += 31;
            int end = html.indexOf("</span>", start);
            if (end > start) {
                return html.substring(start, end).trim();
            }
        }
        return "Statistic";
    }
    
    private Color extractColor(String html) {
        int colorIndex = html.indexOf("color: #");
        if (colorIndex != -1) {
            String hex = html.substring(colorIndex + 8, colorIndex + 14);
            try {
                return Color.decode("#" + hex);
            } catch (Exception e) {
                return Constants.PRIMARY_COLOR;
            }
        }
        return Constants.PRIMARY_COLOR;
    }
    
    private String getMonthName(int month) {
        String[] monthNames = {"January", "February", "March", "April", "May", "June",
                              "July", "August", "September", "October", "November", "December"};
        return month >= 1 && month <= 12 ? monthNames[month - 1] : "Unknown";
    }
}