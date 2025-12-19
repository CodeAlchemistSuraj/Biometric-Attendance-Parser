package org.bioparse.cleaning.ui;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.time.YearMonth;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.awt.event.ActionEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.DocumentEvent;

public class ConfigWizardPanel extends JPanel {
    
    private static final long serialVersionUID = 1L;
    
    // File selection components
    private JTextField inputFileField, mergedFileField, reportFileField;
    private JButton selectInputBtn, selectMergedBtn, selectReportBtn;
    
    // Date selection components
    private JComboBox<Integer> yearComboBox;
    private JComboBox<String> monthComboBox;
    private JLabel currentDateLabel;
    
    // Holiday management components
    private Set<LocalDate> selectedHolidays = new HashSet<>();
    private JPanel calendarPanel;
    private JTextArea manualHolidaysTextArea;
    private JLabel holidaysCountLabel;
    
    // Action buttons
    private JButton processBtn;
    private JButton exportCsvBtn;
    private JButton exportMergedBtn;
    private JLabel statusLabel;
    
    
 // Add these to the component declarations (around line 40)
    private JProgressBar progressBar;
    private JPanel loadingPanel;
    private Timer progressTimer;
    
    
    // Color scheme
    private static final Color PRIMARY_COLOR = new Color(41, 128, 185);
    private static final Color SECONDARY_COLOR = new Color(52, 152, 219);
    private static final Color SUCCESS_COLOR = new Color(39, 174, 96);
    private static final Color WARNING_COLOR = new Color(243, 156, 18);
    private static final Color DANGER_COLOR = new Color(231, 76, 60);
    private static final Color BACKGROUND_COLOR = new Color(248, 249, 252);
    private static final Color CARD_BACKGROUND = Color.WHITE;
    private static final Color TEXT_PRIMARY = new Color(33, 37, 41);
    private static final Color TEXT_SECONDARY = new Color(108, 117, 125);
    private static final Color BORDER_COLOR = new Color(222, 226, 230);
    private static final Color GRID_GAP_COLOR = BACKGROUND_COLOR;
    private static final Color HOLIDAY_SELECTED = new Color(231, 76, 60, 50);
    private static final Color WEEKEND_COLOR = new Color(241, 196, 15, 30);
    private static final Color CALENDAR_HEADER = new Color(248, 249, 250);
    
    // Fonts
    private static final Font FONT_TITLE = new Font("Segoe UI", Font.BOLD, 18);
    private static final Font FONT_SUBTITLE = new Font("Segoe UI", Font.PLAIN, 12);
    private static final Font FONT_HEADER = new Font("Segoe UI", Font.BOLD, 14);
    private static final Font FONT_BODY = new Font("Segoe UI", Font.PLAIN, 13);
    private static final Font FONT_SMALL = new Font("Segoe UI", Font.PLAIN, 11);
    private static final Font FONT_BUTTON = new Font("Segoe UI", Font.BOLD, 13);
    private static final Font FONT_STATUS = new Font("Segoe UI", Font.PLAIN, 12);
    private static final Font FONT_CALENDAR = new Font("Segoe UI", Font.PLAIN, 12);
    private static final Font FONT_CALENDAR_HEADER = new Font("Segoe UI", Font.BOLD, 12);
    
    // Dimensions
    private static final int LABEL_WIDTH = 130;
    private static final int FIELD_HEIGHT = 40;
    private static final int CARD_PADDING = 20;
    private static final int GRID_GAP = 20;
    private static final int ROW_GAP = 12;
    private static final int BUTTON_HEIGHT = 45;
    
    public ConfigWizardPanel(ActionListener processListener, ActionListener exportCsvListener, 
                            ActionListener exportMergedListener) {
        setLayout(new BorderLayout());
        setBackground(BACKGROUND_COLOR);
        
        setupGridContent(processListener, exportCsvListener, exportMergedListener);
        setupFileBrowseActions();
    }
    
    private void setupFileBrowseActions() {
        // Add action listeners for browse buttons
        selectInputBtn.addActionListener(e -> browseForFile(inputFileField, "Select Input Attendance File", "Excel Files", "xlsx", "xls"));
        selectMergedBtn.addActionListener(e -> browseForFile(mergedFileField, "Select Merged Output File", "Excel Files", "xlsx", "xls"));
        selectReportBtn.addActionListener(e -> browseForFile(reportFileField, "Select Report Output File", "Excel Files", "xlsx", "xls"));
    }
    
    private void browseForFile(JTextField field, String title, String description, String... extensions) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(title);
        
        // Set file filter
        if (extensions.length > 0) {
            javax.swing.filechooser.FileNameExtensionFilter filter = 
                new javax.swing.filechooser.FileNameExtensionFilter(description, extensions);
            fileChooser.setFileFilter(filter);
        }
        
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            field.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
    }
    
    private void setupGridContent(ActionListener processListener, ActionListener exportCsvListener,
                                 ActionListener exportMergedListener) {
        // Main grid container - 2x2 grid
        JPanel gridPanel = new JPanel(new GridBagLayout());
        gridPanel.setBackground(GRID_GAP_COLOR);
        gridPanel.setBorder(new EmptyBorder(GRID_GAP, GRID_GAP, GRID_GAP, GRID_GAP));
        
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.insets = new Insets(0, 0, GRID_GAP, GRID_GAP);
        
        // Row 1, Column 1: File Inputs Card
        gbc.gridx = 0;
        gbc.gridy = 0;
        gridPanel.add(createFileInputsCard(), gbc);
        
        // Row 1, Column 2: Time Period Card
        gbc.gridx = 1;
        gbc.gridy = 0;
        gridPanel.add(createTimePeriodCard(), gbc);
        
        // Row 2, Column 1: Holidays Card
        gbc.gridx = 0;
        gbc.gridy = 1;
        gridPanel.add(createHolidaysCard(), gbc);
        
        // Row 2, Column 2: Actions Card
        gbc.gridx = 1;
        gbc.gridy = 1;
        gbc.insets = new Insets(0, 0, 0, 0);
        gridPanel.add(createActionsCard(processListener, exportCsvListener, exportMergedListener), gbc);
        
        add(gridPanel, BorderLayout.CENTER);
        
        // Set minimum size to prevent shrinking
        setMinimumSize(new Dimension(1000, 800));
        setPreferredSize(new Dimension(1200, 850));
    }
    
    // ==================== CARD 1: FILE INPUTS ====================
    
    private JPanel createFileInputsCard() {
        JPanel card = createCard("📁 File Inputs", "Configure input and output file paths");
        
        JPanel content = new JPanel();
        content.setLayout(new BoxLayout(content, BoxLayout.Y_AXIS));
        content.setBackground(CARD_BACKGROUND);
        content.setBorder(new EmptyBorder(5, 0, 0, 0));
        
        // Input Attendance File
        content.add(createCompactFileRow("Input Attendance:", 
            inputFileField = createCompactTextField(""),
            selectInputBtn = createCompactButton("Browse")));
        content.add(Box.createVerticalStrut(ROW_GAP));
        
        // Merged Output File
        content.add(createCompactFileRow("Merged Output:", 
            mergedFileField = createCompactTextField("Merged_Output.xlsx"),
            selectMergedBtn = createCompactButton("Browse")));
        content.add(Box.createVerticalStrut(ROW_GAP));
        
        // Final Report File
        content.add(createCompactFileRow("Final Report:", 
            reportFileField = createCompactTextField("Employee_Report.xlsx"),
            selectReportBtn = createCompactButton("Browse")));
        
        card.add(content, BorderLayout.CENTER);
        return card;
    }
    
    private JPanel createCompactFileRow(String label, JTextField field, JButton button) {
        JPanel row = new JPanel(new BorderLayout(10, 0));
        row.setBackground(CARD_BACKGROUND);
        row.setMaximumSize(new Dimension(Integer.MAX_VALUE, FIELD_HEIGHT));
        
        // Label with fixed width
        JLabel titleLabel = new JLabel(label);
        titleLabel.setFont(FONT_BODY);
        titleLabel.setForeground(TEXT_PRIMARY);
        titleLabel.setPreferredSize(new Dimension(LABEL_WIDTH, FIELD_HEIGHT));
        
        // Text field
        field.setMaximumSize(new Dimension(Integer.MAX_VALUE, FIELD_HEIGHT));
        
        // Button with fixed width
        button.setPreferredSize(new Dimension(80, FIELD_HEIGHT));
        
        row.add(titleLabel, BorderLayout.WEST);
        row.add(field, BorderLayout.CENTER);
        row.add(button, BorderLayout.EAST);
        
        return row;
    }
    
    // ==================== CARD 2: TIME PERIOD ====================
    
    private JPanel createTimePeriodCard() {
        JPanel card = createCard("📅 Time Period", "Select month and year for processing");
        
        JPanel content = new JPanel(new GridBagLayout());
        content.setBackground(CARD_BACKGROUND);
        content.setBorder(new EmptyBorder(20, 0, 20, 0));
        
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(10, 40, 10, 40);
        
        // Year selector
        gbc.gridx = 0;
        gbc.gridy = 0;
        content.add(createCenteredFormField("Year", yearComboBox = createYearComboBox()), gbc);
        
        // Month selector
        gbc.gridy = 1;
        content.add(createCenteredFormField("Month", monthComboBox = createMonthComboBox()), gbc);
        
        // Current date button
        gbc.gridy = 2;
        gbc.insets = new Insets(20, 40, 0, 40);
        content.add(createCurrentDateButton(), gbc);
        
        card.add(content, BorderLayout.CENTER);
        return card;
    }
    
    private JPanel createCenteredFormField(String label, JComponent component) {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(CARD_BACKGROUND);
        
        JLabel fieldLabel = new JLabel(label, SwingConstants.CENTER);
        fieldLabel.setFont(FONT_BODY);
        fieldLabel.setForeground(TEXT_PRIMARY);
        
        // Center the component
        JPanel centerPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 0, 0));
        centerPanel.setBackground(CARD_BACKGROUND);
        centerPanel.add(component);
        
        panel.add(fieldLabel, BorderLayout.NORTH);
        panel.add(centerPanel, BorderLayout.CENTER);
        
        return panel;
    }
    
    private JPanel createCurrentDateButton() {
        JPanel panel = new JPanel(new BorderLayout(10, 0));
        panel.setBackground(CARD_BACKGROUND);
        
        JButton button = new JButton("📅 Use Current Date");
        button.setFont(FONT_SMALL);
        button.setBackground(new Color(248, 249, 250));
        button.setForeground(TEXT_PRIMARY);
        button.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(BORDER_COLOR, 1),
            new EmptyBorder(8, 15, 8, 15)
        ));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setFocusPainted(false);
        button.addActionListener(e -> setCurrentDate());
        
        // Current date display
        currentDateLabel = new JLabel("Today: " + LocalDate.now().format(DateTimeFormatter.ofPattern("MMM d, yyyy")));
        currentDateLabel.setFont(FONT_SMALL);
        currentDateLabel.setForeground(TEXT_SECONDARY);
        currentDateLabel.setHorizontalAlignment(SwingConstants.CENTER);
        
        panel.add(button, BorderLayout.NORTH);
        panel.add(Box.createVerticalStrut(10), BorderLayout.CENTER);
        panel.add(currentDateLabel, BorderLayout.SOUTH);
        
        return panel;
    }
    
    // ==================== CARD 3: HOLIDAYS ====================
    
    private JPanel createHolidaysCard() {
        JPanel card = createCard("🗓️ Holidays & Weekends", "Select non-working days or manually enter dates");
        
        JPanel content = new JPanel(new BorderLayout(0, 10));
        content.setBackground(CARD_BACKGROUND);
        content.setBorder(new EmptyBorder(5, 0, 0, 0));
        
        // Tabbed pane for calendar and manual entry
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setFont(FONT_BODY);
        
        // Tab 1: Calendar View
        JPanel calendarTab = new JPanel(new BorderLayout(0, 10));
        calendarTab.setBackground(CARD_BACKGROUND);
        
        // Month navigation
        JPanel navPanel = new JPanel(new BorderLayout(10, 0));
        navPanel.setBackground(CARD_BACKGROUND);
        navPanel.setBorder(new EmptyBorder(0, 0, 10, 0));
        
        JButton prevMonthBtn = new JButton("◀");
        prevMonthBtn.setFont(FONT_BODY);
        prevMonthBtn.addActionListener(e -> navigateMonth(-1));
        
        JButton nextMonthBtn = new JButton("▶");
        nextMonthBtn.setFont(FONT_BODY);
        nextMonthBtn.addActionListener(e -> navigateMonth(1));
        
        navPanel.add(prevMonthBtn, BorderLayout.WEST);
        navPanel.add(nextMonthBtn, BorderLayout.EAST);
        
        // Calendar panel
        calendarPanel = new JPanel(new GridLayout(0, 7, 2, 2));
        calendarPanel.setBackground(CARD_BACKGROUND);
        
        calendarTab.add(navPanel, BorderLayout.NORTH);
        calendarTab.add(calendarPanel, BorderLayout.CENTER);
        
        // Tab 2: Manual Entry
        JPanel manualTab = new JPanel(new BorderLayout(0, 10));
        manualTab.setBackground(CARD_BACKGROUND);
        
        JLabel manualLabel = new JLabel("<html>Enter dates (comma or newline separated):<br><i>Example: 1, 15, 25 or one per line</i></html>");
        manualLabel.setFont(FONT_SMALL);
        manualLabel.setForeground(TEXT_SECONDARY);
        manualLabel.setBorder(new EmptyBorder(0, 0, 5, 0));
        
        manualHolidaysTextArea = new JTextArea();
        manualHolidaysTextArea.setFont(FONT_BODY);
        manualHolidaysTextArea.setLineWrap(true);
        manualHolidaysTextArea.setWrapStyleWord(true);
        manualHolidaysTextArea.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(BORDER_COLOR, 1),
            new EmptyBorder(10, 15, 10, 15)
        ));
        
        // Add listener to sync manual entries with calendar
        manualHolidaysTextArea.getDocument().addDocumentListener(new DocumentListener() {
            @Override
            public void insertUpdate(DocumentEvent e) {
                syncManualHolidays();
            }
            @Override
            public void removeUpdate(DocumentEvent e) {
                syncManualHolidays();
            }
            @Override
            public void changedUpdate(DocumentEvent e) {
                syncManualHolidays();
            }
        });
        
        JScrollPane manualScroll = new JScrollPane(manualHolidaysTextArea);
        manualScroll.setBorder(null);
        
        manualTab.add(manualLabel, BorderLayout.NORTH);
        manualTab.add(manualScroll, BorderLayout.CENTER);
        
        // Add tabs
        tabbedPane.addTab("Calendar", calendarTab);
        tabbedPane.addTab("Manual Entry", manualTab);
        
        // Holidays count and action buttons
        JPanel bottomPanel = new JPanel(new BorderLayout(10, 0));
        bottomPanel.setBackground(CARD_BACKGROUND);
        bottomPanel.setBorder(new EmptyBorder(10, 0, 0, 0));
        
        holidaysCountLabel = new JLabel("0 holidays selected");
        holidaysCountLabel.setFont(FONT_SMALL);
        holidaysCountLabel.setForeground(TEXT_SECONDARY);
        
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 5, 0));
        buttonPanel.setBackground(CARD_BACKGROUND);
        
        JButton clearBtn = new JButton("Clear All");
        clearBtn.setFont(FONT_SMALL);
        clearBtn.addActionListener(e -> clearAllHolidays());
        
        buttonPanel.add(clearBtn);
        
        bottomPanel.add(holidaysCountLabel, BorderLayout.WEST);
        bottomPanel.add(buttonPanel, BorderLayout.EAST);
        
        content.add(tabbedPane, BorderLayout.CENTER);
        content.add(bottomPanel, BorderLayout.SOUTH);
        
        card.add(content, BorderLayout.CENTER);
        
        // Initialize calendar AFTER creating the count label
        updateCalendar();
        
        return card;
    }
    
    private void updateCalendar() {
        if (calendarPanel == null) return;
        
        calendarPanel.removeAll();
        
        int year = getSelectedYear();
        int month = getSelectedMonth();
        LocalDate firstDay = LocalDate.of(year, month, 1);
        int daysInMonth = YearMonth.of(year, month).lengthOfMonth();
        
        // Day headers
        String[] dayNames = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};
        for (String dayName : dayNames) {
            JLabel header = new JLabel(dayName, SwingConstants.CENTER);
            header.setFont(FONT_CALENDAR_HEADER);
            header.setForeground(TEXT_PRIMARY);
            header.setBackground(CALENDAR_HEADER);
            header.setOpaque(true);
            header.setBorder(BorderFactory.createLineBorder(BORDER_COLOR, 1));
            calendarPanel.add(header);
        }
        
        // Empty cells for days before the first day of month
        int firstDayOfWeek = firstDay.getDayOfWeek().getValue() % 7; // Sunday = 0
        for (int i = 0; i < firstDayOfWeek; i++) {
            calendarPanel.add(new JLabel(""));
        }
        
        // Day cells
        for (int day = 1; day <= daysInMonth; day++) {
            LocalDate currentDate = LocalDate.of(year, month, day);
            JButton dayButton = new JButton(String.valueOf(day));
            dayButton.setFont(FONT_CALENDAR);
            dayButton.setFocusPainted(false);
            dayButton.setCursor(new Cursor(Cursor.HAND_CURSOR));
            
            // Check if it's a weekend
            if (currentDate.getDayOfWeek().getValue() >= 6) { // Saturday or Sunday
                dayButton.setBackground(WEEKEND_COLOR);
            } else {
                dayButton.setBackground(Color.WHITE);
            }
            
            // Check if it's selected as holiday
            if (selectedHolidays.contains(currentDate)) {
                dayButton.setBackground(HOLIDAY_SELECTED);
                dayButton.setForeground(Color.RED);
            } else {
                dayButton.setForeground(TEXT_PRIMARY);
            }
            
            dayButton.setBorder(BorderFactory.createLineBorder(BORDER_COLOR, 1));
            
            // Add action listener
            dayButton.addActionListener(e -> toggleHoliday(currentDate, dayButton));
            
            calendarPanel.add(dayButton);
        }
        
        calendarPanel.revalidate();
        calendarPanel.repaint();
        
        // Only update count if label exists
        if (holidaysCountLabel != null) {
            updateHolidaysCount();
        }
    }
    
    private void toggleHoliday(LocalDate date, JButton button) {
        if (selectedHolidays.contains(date)) {
            selectedHolidays.remove(date);
            if (date.getDayOfWeek().getValue() >= 6) {
                button.setBackground(WEEKEND_COLOR);
            } else {
                button.setBackground(Color.WHITE);
            }
            button.setForeground(TEXT_PRIMARY);
        } else {
            selectedHolidays.add(date);
            button.setBackground(HOLIDAY_SELECTED);
            button.setForeground(Color.RED);
        }
        updateManualHolidaysText();
        updateHolidaysCount();
    }
    
    private void navigateMonth(int direction) {
        int currentYear = getSelectedYear();
        int currentMonth = getSelectedMonth();
        
        if (direction > 0) {
            if (currentMonth == 12) {
                currentYear++;
                currentMonth = 1;
            } else {
                currentMonth++;
            }
        } else {
            if (currentMonth == 1) {
                currentYear--;
                currentMonth = 12;
            } else {
                currentMonth--;
            }
        }
        
        yearComboBox.setSelectedItem(currentYear);
        monthComboBox.setSelectedIndex(currentMonth - 1);
        updateCalendar();
    }
    
    private void updateManualHolidaysText() {
        StringBuilder sb = new StringBuilder();
        for (LocalDate date : selectedHolidays) {
            sb.append(date.getDayOfMonth()).append(", ");
        }
        if (sb.length() > 0) {
            sb.setLength(sb.length() - 2); // Remove trailing comma and space
        }
        manualHolidaysTextArea.setText(sb.toString());
    }
    
    private void syncManualHolidays() {
        selectedHolidays.clear();
        String text = manualHolidaysTextArea.getText().trim();
        
        if (text.isEmpty()) {
            updateCalendar();
            return;
        }
        
        String[] parts = text.split("[,\\n]");
        for (String part : parts) {
            try {
                String cleanPart = part.replaceAll("[^0-9]", "").trim();
                if (!cleanPart.isEmpty()) {
                    int day = Integer.parseInt(cleanPart);
                    int year = getSelectedYear();
                    int month = getSelectedMonth();
                    
                    try {
                        LocalDate date = LocalDate.of(year, month, day);
                        selectedHolidays.add(date);
                    } catch (Exception e) {
                        // Invalid date for this month
                    }
                }
            } catch (NumberFormatException e) {
                // Skip non-numeric entries
            }
        }
        
        updateCalendar();
    }
    
    private void clearAllHolidays() {
        int result = JOptionPane.showConfirmDialog(this,
            "Are you sure you want to clear all selected holidays?",
            "Confirm Clear",
            JOptionPane.YES_NO_OPTION,
            JOptionPane.WARNING_MESSAGE);
            
        if (result == JOptionPane.YES_OPTION) {
            selectedHolidays.clear();
            manualHolidaysTextArea.setText("");
            updateCalendar();
            setStatus("All holidays cleared", SUCCESS_COLOR);
        }
    }
    
    private void updateHolidaysCount() {
        if (holidaysCountLabel != null) {
            int count = selectedHolidays.size();
            holidaysCountLabel.setText(count + " holiday" + (count != 1 ? "s" : "") + " selected");
        }
    }
    
    // ==================== CARD 4: ACTIONS ====================
    
private JPanel createActionsCard(ActionListener processListener, ActionListener exportCsvListener,
                                ActionListener exportMergedListener) {
    JPanel card = createCard("▶️ Actions & Status", "Process data and export reports");
    
    JPanel content = new JPanel(new BorderLayout(0, 20));
    content.setBackground(CARD_BACKGROUND);
    content.setBorder(new EmptyBorder(10, 0, 0, 0));
    
    // Button panel
    JPanel buttonPanel = new JPanel();
    buttonPanel.setLayout(new BoxLayout(buttonPanel, BoxLayout.Y_AXIS));
    buttonPanel.setBackground(CARD_BACKGROUND);
    
    // Merge button
    processBtn = createActionButton("🔧 Merge & Process Data", PRIMARY_COLOR, 
                                   "Process attendance data and generate report", processListener);
    processBtn.setAlignmentX(Component.CENTER_ALIGNMENT);
    processBtn.setMaximumSize(new Dimension(Integer.MAX_VALUE, BUTTON_HEIGHT));
    buttonPanel.add(processBtn);
    buttonPanel.add(Box.createVerticalStrut(15));
    
    // Export buttons panel
    JPanel exportPanel = new JPanel();
    exportPanel.setLayout(new BoxLayout(exportPanel, BoxLayout.Y_AXIS));
    exportPanel.setBackground(CARD_BACKGROUND);
    exportPanel.setAlignmentX(Component.CENTER_ALIGNMENT);
    
    exportCsvBtn = createActionButton("📊 Export CSV Report", SUCCESS_COLOR,
                                     "Export report to CSV format", exportCsvListener);
    exportCsvBtn.setEnabled(false);
    exportCsvBtn.setMaximumSize(new Dimension(Integer.MAX_VALUE, BUTTON_HEIGHT));
    exportCsvBtn.setAlignmentX(Component.CENTER_ALIGNMENT);
    
    exportMergedBtn = createActionButton("📁 Export Merged File", new Color(155, 89, 182),
                                       "Export merged attendance file", exportMergedListener);
    exportMergedBtn.setEnabled(false);
    exportMergedBtn.setMaximumSize(new Dimension(Integer.MAX_VALUE, BUTTON_HEIGHT));
    exportMergedBtn.setAlignmentX(Component.CENTER_ALIGNMENT);
    
    exportPanel.add(exportCsvBtn);
    exportPanel.add(Box.createVerticalStrut(10));
    exportPanel.add(exportMergedBtn);
    
    buttonPanel.add(exportPanel);
    
    // Loading panel (initially hidden)
    loadingPanel = new JPanel();
    loadingPanel.setLayout(new BoxLayout(loadingPanel, BoxLayout.Y_AXIS));
    loadingPanel.setBackground(CARD_BACKGROUND);
    loadingPanel.setAlignmentX(Component.CENTER_ALIGNMENT);
    loadingPanel.setVisible(false);
    
    // Spinner with text
    JPanel spinnerPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 0));
    spinnerPanel.setBackground(CARD_BACKGROUND);
    
    JLabel spinnerLabel = new JLabel("⏳");
    spinnerLabel.setFont(new Font("Segoe UI", Font.PLAIN, 20));
    
    JLabel loadingText = new JLabel("Processing data...");
    loadingText.setFont(FONT_BODY);
    loadingText.setForeground(PRIMARY_COLOR);
    
    spinnerPanel.add(spinnerLabel);
    spinnerPanel.add(loadingText);
    
    // Progress bar
    progressBar = new JProgressBar();
    progressBar.setIndeterminate(true); // Makes it a spinner animation
    progressBar.setPreferredSize(new Dimension(200, 20));
    progressBar.setVisible(false);
    
    // Optional: determinate progress bar for longer operations
    JProgressBar determinateProgress = new JProgressBar(0, 100);
    determinateProgress.setPreferredSize(new Dimension(200, 20));
    determinateProgress.setVisible(false);
    determinateProgress.setStringPainted(true);
    
    loadingPanel.add(spinnerPanel);
    loadingPanel.add(Box.createVerticalStrut(10));
    loadingPanel.add(progressBar);
    loadingPanel.add(Box.createVerticalStrut(5));
    loadingPanel.add(determinateProgress);
    
    // Add loading panel to button panel
    buttonPanel.add(Box.createVerticalStrut(10));
    buttonPanel.add(loadingPanel);
    
    // Status panel
    JPanel statusPanel = new JPanel(new BorderLayout());
    statusPanel.setBackground(CARD_BACKGROUND);
    statusPanel.setBorder(new EmptyBorder(20, 0, 0, 0));
    
    statusLabel = new JLabel("Ready to configure...");
    statusLabel.setFont(FONT_STATUS);
    statusLabel.setForeground(TEXT_SECONDARY);
    statusLabel.setHorizontalAlignment(SwingConstants.CENTER);
    statusLabel.setBorder(BorderFactory.createCompoundBorder(
        BorderFactory.createMatteBorder(1, 0, 0, 0, BORDER_COLOR),
        new EmptyBorder(10, 0, 0, 0)
    ));
    
    statusPanel.add(statusLabel, BorderLayout.CENTER);
    
    content.add(buttonPanel, BorderLayout.CENTER);
    content.add(statusPanel, BorderLayout.SOUTH);
    
    card.add(content, BorderLayout.CENTER);
    return card;
} 
    // ==================== HELPER METHODS ====================
    
    private JPanel createCard(String title, String subtitle) {
        JPanel card = new JPanel(new BorderLayout());
        card.setBackground(CARD_BACKGROUND);
        card.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(BORDER_COLOR, 1),
            new EmptyBorder(CARD_PADDING, CARD_PADDING, CARD_PADDING, CARD_PADDING)
        ));
        
        // Card header
        JPanel header = new JPanel(new BorderLayout());
        header.setBackground(CARD_BACKGROUND);
        header.setBorder(new EmptyBorder(0, 0, 15, 0));
        
        JLabel titleLabel = new JLabel(title);
        titleLabel.setFont(FONT_HEADER);
        titleLabel.setForeground(TEXT_PRIMARY);
        
        JLabel subtitleLabel = new JLabel(subtitle);
        subtitleLabel.setFont(FONT_SMALL);
        subtitleLabel.setForeground(TEXT_SECONDARY);
        
        header.add(titleLabel, BorderLayout.NORTH);
        header.add(subtitleLabel, BorderLayout.CENTER);
        
        card.add(header, BorderLayout.NORTH);
        return card;
    }
    
    
    
    
    
    
    
    private JTextField createCompactTextField(String placeholder) {
        JTextField field = new JTextField(placeholder);
        field.setFont(FONT_BODY);
        field.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(BORDER_COLOR, 1),
            new EmptyBorder(0, 12, 0, 12)
        ));
        field.setBackground(Color.WHITE);
        field.setPreferredSize(new Dimension(200, FIELD_HEIGHT));
        return field;
    }
    
    private JButton createCompactButton(String text) {
        JButton button = new JButton(text);
        button.setFont(FONT_SMALL);
        button.setBackground(PRIMARY_COLOR);
        button.setForeground(Color.WHITE);
        button.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(PRIMARY_COLOR.darker(), 1),
            new EmptyBorder(8, 15, 8, 15)
        ));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setFocusPainted(false);
        
        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                button.setBackground(SECONDARY_COLOR);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                button.setBackground(PRIMARY_COLOR);
            }
        });
        
        return button;
    }
    
    private JComboBox<Integer> createYearComboBox() {
        JComboBox<Integer> combo = new JComboBox<>();
        combo.setFont(FONT_BODY);
        combo.setBackground(Color.WHITE);
        combo.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(BORDER_COLOR, 1),
            new EmptyBorder(0, 12, 0, 12)
        ));
        combo.setPreferredSize(new Dimension(150, FIELD_HEIGHT));
        
        int currentYear = LocalDate.now().getYear();
        for (int year = currentYear - 2; year <= currentYear + 2; year++) {
            combo.addItem(year);
        }
        combo.setSelectedItem(currentYear);
        
        // Update calendar when year changes
        combo.addActionListener(e -> updateCalendar());
        
        return combo;
    }
    
    private JComboBox<String> createMonthComboBox() {
        JComboBox<String> combo = new JComboBox<>();
        combo.setFont(FONT_BODY);
        combo.setBackground(Color.WHITE);
        combo.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(BORDER_COLOR, 1),
            new EmptyBorder(0, 12, 0, 12)
        ));
        combo.setPreferredSize(new Dimension(150, FIELD_HEIGHT));
        
        String[] months = {"January", "February", "March", "April", "May", "June", 
                          "July", "August", "September", "October", "November", "December"};
        for (String month : months) {
            combo.addItem(month);
        }
        combo.setSelectedIndex(LocalDate.now().getMonthValue() - 1);
        
        // Update calendar when month changes
        combo.addActionListener(e -> updateCalendar());
        
        return combo;
    }
    
    private JButton createActionButton(String text, Color color, String tooltip, ActionListener listener) {
        JButton button = new JButton(text);
        button.setFont(FONT_BUTTON);
        button.setBackground(color);
        button.setForeground(Color.WHITE);
        button.setToolTipText(tooltip);
        button.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(color.darker(), 1),
            new EmptyBorder(12, 20, 12, 20)
        ));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setFocusPainted(false);
        button.addActionListener(listener);
        
        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                button.setBackground(color.brighter());
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                button.setBackground(color);
            }
        });
        
        return button;
    }
    
    
    
    
    
 // Add these methods to ConfigWizardPanel.java

    public void showLoading() {
        SwingUtilities.invokeLater(() -> {
            processBtn.setEnabled(false);
            processBtn.setText("⏳ Processing...");
            loadingPanel.setVisible(true);
            progressBar.setVisible(true);
            statusLabel.setText("Processing attendance data... This may take a moment.");
            statusLabel.setForeground(PRIMARY_COLOR);
            revalidate();
            repaint();
        });
    }

    public void hideLoading() {
        SwingUtilities.invokeLater(() -> {
            processBtn.setEnabled(true);
            processBtn.setText("🔧 Merge & Process Data");
            loadingPanel.setVisible(false);
            progressBar.setVisible(false);
            revalidate();
            repaint();
        });
    }

    public void showProgress(int progress, String message) {
        SwingUtilities.invokeLater(() -> {
            statusLabel.setText(message + " (" + progress + "%)");
        });
    }
    
    
    
    // ==================== BUSINESS LOGIC ====================
    
    private void setCurrentDate() {
        LocalDate now = LocalDate.now();
        yearComboBox.setSelectedItem(now.getYear());
        monthComboBox.setSelectedIndex(now.getMonthValue() - 1);
        
        currentDateLabel.setText("Applied: " + now.format(DateTimeFormatter.ofPattern("MMM d, yyyy")));
        setStatus("Current date applied successfully", SUCCESS_COLOR);
        updateCalendar();
    }
    
    private void setStatus(String message, Color color) {
        if (statusLabel != null) {
            statusLabel.setText(message);
            statusLabel.setForeground(color);
        }
    }
    
    // ==================== PUBLIC GETTERS ====================
    
    public String getInputFilePath() {
        return inputFileField.getText().trim();
    }
    
    public String getMergedFilePath() {
        return mergedFileField.getText().trim();
    }
    
    public String getReportFilePath() {
        return reportFileField.getText().trim();
    }
    
    public int getSelectedYear() {
        return (Integer) yearComboBox.getSelectedItem();
    }
    
    public int getSelectedMonth() {
        return getMonthNumber((String) monthComboBox.getSelectedItem());
    }
    
    public List<Integer> getHolidays() {
        List<Integer> holidays = new ArrayList<>();
        for (LocalDate date : selectedHolidays) {
            holidays.add(date.getDayOfMonth());
        }
        return holidays;
    }
    
    public Set<LocalDate> getHolidayDates() {
        return new HashSet<>(selectedHolidays);
    }
    
    public void enableExportButtons(boolean enabled) {
        if (exportCsvBtn != null && exportMergedBtn != null) {
            exportCsvBtn.setEnabled(enabled);
            exportMergedBtn.setEnabled(enabled);
            
            if (enabled) {
                exportCsvBtn.setBackground(SUCCESS_COLOR);
                exportMergedBtn.setBackground(new Color(155, 89, 182));
                setStatus("Processing complete. Ready to export.", SUCCESS_COLOR);
            } else {
                exportCsvBtn.setBackground(new Color(189, 195, 199));
                exportMergedBtn.setBackground(new Color(189, 195, 199));
            }
        }
    }
    
    public void setStatusMessage(String message, boolean isError) {
        setStatus(message, isError ? DANGER_COLOR : SUCCESS_COLOR);
    }
    
    private int getMonthNumber(String monthName) {
        String[] months = {"January", "February", "March", "April", "May", "June", 
                          "July", "August", "September", "October", "November", "December"};
        for (int i = 0; i < months.length; i++) {
            if (months[i].equalsIgnoreCase(monthName)) {
                return i + 1;
            }
        }
        return 1;
    }
}