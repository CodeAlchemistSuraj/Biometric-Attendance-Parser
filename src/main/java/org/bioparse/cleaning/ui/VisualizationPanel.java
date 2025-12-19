package org.bioparse.cleaning.ui;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;

import org.bioparse.cleaning.AttendanceMerger;
import org.bioparse.cleaning.ChartDataGenerator;

public class VisualizationPanel extends JPanel {
    
    private static final long serialVersionUID = 1L;
    private JPanel contentPanel;
    private AttendanceMerger.MergeResult mergeResult;
    
    public VisualizationPanel() {
        setLayout(new BorderLayout());
        setBackground(new Color(242, 244, 248));
        
        setupHeader();
        setupContent();
    }
    
    private void setupHeader() {
        JPanel header = new JPanel(new BorderLayout());
        header.setBackground(new Color(255, 255, 255));
        header.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createMatteBorder(0, 0, 1, 0, new Color(225, 228, 232)),
            BorderFactory.createEmptyBorder(12, 20, 12, 20)
        ));
        
        JLabel title = new JLabel("📊 Analytics Dashboard");
        title.setFont(new Font("Segoe UI Semibold", Font.PLAIN, 18));
        title.setForeground(new Color(52, 58, 64));
        
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        buttonPanel.setOpaque(false);
        
        JButton refreshBtn = createCompactButton("⟳ Refresh", new Color(66, 133, 244));
        refreshBtn.addActionListener(e -> refreshCharts());
        
        buttonPanel.add(refreshBtn);
        
        header.add(title, BorderLayout.WEST);
        header.add(buttonPanel, BorderLayout.EAST);
        add(header, BorderLayout.NORTH);
    }
    
    private void setupContent() {
        contentPanel = new JPanel(new BorderLayout());
        contentPanel.setBackground(new Color(242, 244, 248));
        contentPanel.setBorder(new EmptyBorder(15, 15, 15, 15));
        contentPanel.add(createWelcomePanel(), BorderLayout.CENTER);
        
        JScrollPane scrollPane = new JScrollPane(contentPanel);
        scrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        scrollPane.getVerticalScrollBar().setUnitIncrement(12);
        scrollPane.setBorder(null);
        scrollPane.getViewport().setBackground(new Color(242, 244, 248));
        
        add(scrollPane, BorderLayout.CENTER);
    }
    
    private JPanel createWelcomePanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(new Color(242, 244, 248));
        
        JLabel welcome = new JLabel(
            "<html><div style='text-align: center; padding: 40px 20px;'>" +
            "<div style='font-size: 24px; color: #495057; margin-bottom: 12px;'>📈 Welcome</div>" +
            "<div style='font-size: 14px; color: #6c757d; max-width: 500px; margin: 0 auto;'>" +
            "Process attendance data to view analytics, trends, and insights.</div>" +
            "</div></html>", SwingConstants.CENTER);
        
        panel.add(welcome, BorderLayout.CENTER);
        return panel;
    }
    
    private JButton createCompactButton(String text, Color bgColor) {
        JButton btn = new JButton(text);
        btn.setBackground(bgColor);
        btn.setForeground(Color.WHITE);
        btn.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        btn.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(bgColor.darker(), 1),
            BorderFactory.createEmptyBorder(4, 12, 4, 12)
        ));
        btn.setFocusPainted(false);
        btn.setCursor(new Cursor(Cursor.HAND_CURSOR));
        return btn;
    }
    
    public void setMergeResult(AttendanceMerger.MergeResult mergeResult) {
        this.mergeResult = mergeResult;
        refreshCharts();
    }
    
    public void refreshCharts() {
        contentPanel.removeAll();
        
        if (mergeResult == null || mergeResult.allEmployees.isEmpty()) {
            contentPanel.add(createWelcomePanel(), BorderLayout.CENTER);
        } else {
            contentPanel.add(createDashboardView(), BorderLayout.CENTER);
        }
        
        contentPanel.revalidate();
        contentPanel.repaint();
    }
    
    private JPanel createDashboardView() {
        JPanel dashboard = new JPanel(new GridBagLayout());
        dashboard.setBackground(new Color(242, 244, 248));
        
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(8, 8, 8, 8);
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.gridwidth = 1;
        
        // First Row - Overview Charts
        gbc.gridx = 0; gbc.gridy = 0;
        dashboard.add(createChartCard("Daily Trend", 
            ChartDataGenerator.createAttendanceTrendChart(mergeResult.allEmployees, mergeResult.monthDays)), gbc);
        
        gbc.gridx = 1;
        dashboard.add(createChartCard("Attendance", 
            ChartDataGenerator.createAttendanceDistributionChart(mergeResult.allEmployees, mergeResult.monthDays)), gbc);
        
        gbc.gridx = 2;
        dashboard.add(createChartCard("Departments", 
            ChartDataGenerator.createDepartmentWiseAttendanceChart(mergeResult.allEmployees)), gbc);
        
        // Second Row - Performance Charts
        gbc.gridx = 0; gbc.gridy = 1;
        dashboard.add(createChartCard("Performance", 
            ChartDataGenerator.createTopPerformersChart(mergeResult.allEmployees, mergeResult.monthDays)), gbc);
        
        gbc.gridx = 1;
        dashboard.add(createChartCard("Late Arrivals", 
            ChartDataGenerator.createLateArrivalsAnalysis(mergeResult.allEmployees, mergeResult.monthDays)), gbc);
        
        gbc.gridx = 2;
        dashboard.add(createChartCard("Overtime", 
            ChartDataGenerator.createOvertimeAnalysisChart(mergeResult.allEmployees)), gbc);
        
        // Third Row - Additional Charts
        gbc.gridx = 0; gbc.gridy = 2;
        dashboard.add(createChartCard("Trend Line", 
            ChartDataGenerator.createDailyAttendanceLineChart(mergeResult.allEmployees, mergeResult.monthDays)), gbc);
        
        gbc.gridx = 1;
        dashboard.add(createChartCard("Weekly", 
            ChartDataGenerator.createWeeklyAttendanceChart(mergeResult.allEmployees, mergeResult.monthDays)), gbc);
        
        gbc.gridx = 2;
        dashboard.add(createChartCard("Punctuality", 
            ChartDataGenerator.createShiftComplianceChart(mergeResult.allEmployees)), gbc);
        
        // Add an empty panel to push everything up
        gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 3;
        gbc.weighty = 10.0;
        dashboard.add(Box.createGlue(), gbc);
        
        return dashboard;
    }
    
    private JPanel createChartCard(String title, JPanel chart) {
        JPanel card = new JPanel(new BorderLayout());
        card.setBackground(Color.WHITE);
        card.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(233, 236, 239)),
            BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));
        
        JLabel titleLabel = new JLabel(title);
        titleLabel.setFont(new Font("Segoe UI Semibold", Font.PLAIN, 13));
        titleLabel.setForeground(new Color(73, 80, 87));
        titleLabel.setBorder(new EmptyBorder(0, 0, 8, 0));
        card.add(titleLabel, BorderLayout.NORTH);
        
        chart.setBackground(Color.WHITE);
        card.add(chart, BorderLayout.CENTER);
        
        // Set preferred size for consistency
        card.setPreferredSize(new Dimension(220, 220));
        card.setMaximumSize(new Dimension(240, 240));
        card.setMinimumSize(new Dimension(200, 200));
        
        return card;
    }
}