package org.bioparse.cleaning.ui;

import javax.swing.*;
import javax.swing.border.*;
import javax.swing.table.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

import org.bioparse.cleaning.AttendanceMerger;
import org.bioparse.cleaning.Constants;
import org.bioparse.cleaning.AttendanceQueryViewer;

public class DataViewerPanel extends JPanel {
    
    private static final long serialVersionUID = 1L;
    
    // Modern color palette
    private static final Color CARD_BG = new Color(255, 255, 255);
    private static final Color CARD_SHADOW = new Color(0, 0, 0, 20);
    private static final Color ACCENT_BLUE = new Color(70, 130, 240);
    private static final Color ACCENT_GREEN = new Color(40, 180, 100);
    private static final Color ACCENT_ORANGE = new Color(255, 140, 0);
    private static final Color ACCENT_PURPLE = new Color(155, 89, 182);
    private static final Color LIGHT_BG = new Color(248, 250, 252);
    private static final Color BORDER_COLOR = new Color(230, 234, 240);
    private static final Color TABLE_HEADER_BG = new Color(245, 247, 250);
    private static final Color TABLE_HEADER_TEXT = new Color(60, 70, 85);
    
    private JTextField empCodeField;
    private JSpinner daySpinner, startDaySpinner, endDaySpinner;
    private JButton queryEmpBtn, queryDayBtn, queryRangeBtn, exportBtn, clearBtn;
    private JTable dataTable;
    private DefaultTableModel dataModel;
    private JLabel statusLabel, panelTitleLabel, statsLabel;
    private JPanel mainContentPanel, queryCardsPanel, resultsPanel;
    private JScrollPane mainScrollPane, tableScrollPane;
    private JPanel statsPanel;
    
    private AttendanceMerger.MergeResult mergeResult;
    private AttendanceQueryViewer queryViewer;
    private String mergedFilePath;
    private int currentYear, currentMonth;
    
    public DataViewerPanel(ActionListener queryListener, ActionListener exportListener) {
        setLayout(new BorderLayout());
        setBackground(LIGHT_BG);
        
        setupHeaderPanel();
        setupMainContentPanel(queryListener, exportListener);
        
        // Initialize status
        updateStatus("⏳ Waiting for data... Process attendance data first.");
    }
    
    private void setupHeaderPanel() {
        JPanel headerPanel = new JPanel(new BorderLayout());
        headerPanel.setBackground(new Color(255, 255, 255));
        headerPanel.setBorder(BorderFactory.createCompoundBorder(
            new MatteBorder(0, 0, 1, 0, BORDER_COLOR),
            new EmptyBorder(20, 30, 20, 30)
        ));
        
        // Left side - Title
        JPanel titlePanel = new JPanel(new BorderLayout(10, 0));
        titlePanel.setBackground(Color.WHITE);
        
        JLabel iconLabel = new JLabel("📊");
        iconLabel.setFont(new Font("Segoe UI Emoji", Font.PLAIN, 32));
        
        JPanel textPanel = new JPanel();
        textPanel.setLayout(new BoxLayout(textPanel, BoxLayout.Y_AXIS));
        textPanel.setBackground(Color.WHITE);
        
        panelTitleLabel = new JLabel("Data Query Explorer");
        panelTitleLabel.setFont(new Font("Segoe UI", Font.BOLD, 24));
        panelTitleLabel.setForeground(new Color(30, 40, 55));
        
        JLabel subtitleLabel = new JLabel("Search and analyze attendance records with powerful filters");
        subtitleLabel.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        subtitleLabel.setForeground(new Color(100, 110, 125));
        
        textPanel.add(panelTitleLabel);
        textPanel.add(Box.createVerticalStrut(4));
        textPanel.add(subtitleLabel);
        
        titlePanel.add(iconLabel, BorderLayout.WEST);
        titlePanel.add(textPanel, BorderLayout.CENTER);
        
        // Right side - Stats
        statsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 20, 0));
        statsPanel.setBackground(Color.WHITE);
        
        addStatItem("👥 Employees", "0", ACCENT_BLUE);
        addStatItem("📅 Month Days", "0", ACCENT_GREEN);
        addStatItem("📁 Data Status", "Not Loaded", ACCENT_ORANGE);
        
        headerPanel.add(titlePanel, BorderLayout.WEST);
        headerPanel.add(statsPanel, BorderLayout.EAST);
        
        add(headerPanel, BorderLayout.NORTH);
    }
    
    private void addStatItem(String label, String value, Color color) {
        JPanel statCard = new JPanel();
        statCard.setLayout(new BoxLayout(statCard, BoxLayout.Y_AXIS));
        statCard.setBackground(Color.WHITE);
        statCard.setBorder(new EmptyBorder(5, 15, 5, 15));
        
        JLabel valueLabel = new JLabel(value);
        valueLabel.setFont(new Font("Segoe UI", Font.BOLD, 16));
        valueLabel.setForeground(color);
        valueLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        JLabel labelLabel = new JLabel(label);
        labelLabel.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        labelLabel.setForeground(new Color(120, 130, 145));
        labelLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        statCard.add(valueLabel);
        statCard.add(Box.createVerticalStrut(2));
        statCard.add(labelLabel);
        
        statsPanel.add(statCard);
        
        // Add separator except for last item
        if (statsPanel.getComponentCount() < 3) {
            JSeparator separator = new JSeparator(JSeparator.VERTICAL);
            separator.setPreferredSize(new Dimension(1, 40));
            separator.setBackground(new Color(230, 235, 240));
            statsPanel.add(separator);
        }
    }
    
    private void setupMainContentPanel(ActionListener queryListener, ActionListener exportListener) {
        mainContentPanel = new JPanel();
        mainContentPanel.setLayout(new BoxLayout(mainContentPanel, BoxLayout.Y_AXIS));
        mainContentPanel.setBackground(LIGHT_BG);
        mainContentPanel.setBorder(new EmptyBorder(20, 30, 30, 30));
        
        // Query Cards Section
        setupQueryCardsPanel();
        mainContentPanel.add(queryCardsPanel);
        mainContentPanel.add(Box.createVerticalStrut(25));
        
        // Results Section
        setupResultsPanel();
        mainContentPanel.add(resultsPanel);
        mainContentPanel.add(Box.createVerticalStrut(20));
        
        // Status Bar

        mainContentPanel.add(createStatusBar());
        
        // Wrap in scroll pane
        mainScrollPane = new JScrollPane(mainContentPanel);
        mainScrollPane.setBorder(null);
        mainScrollPane.getViewport().setBackground(LIGHT_BG);
        mainScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        mainScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        
        // Customize scrollbar
        JScrollBar verticalScrollBar = mainScrollPane.getVerticalScrollBar();
        verticalScrollBar.setUnitIncrement(16);
        verticalScrollBar.setBackground(LIGHT_BG);
        verticalScrollBar.setBorder(BorderFactory.createEmptyBorder());
        
        add(mainScrollPane, BorderLayout.CENTER);
    }
    
    private void setupQueryCardsPanel() {
        queryCardsPanel = new JPanel();
        queryCardsPanel.setLayout(new GridLayout(1, 3, 20, 0));
        queryCardsPanel.setBackground(LIGHT_BG);
        queryCardsPanel.setBorder(new EmptyBorder(0, 0, 10, 0));
        
        // Card 1: Employee Query
        queryCardsPanel.add(createQueryCard(
            "👤 Employee Lookup",
            "Search attendance by employee ID",
            ACCENT_BLUE,
            new Component[] {
                createInputField("Employee ID", "Enter employee code...", 200),
                createStyledButton("🔍 Search Employee", ACCENT_BLUE, "queryEmp")
            }
        ));
        
        // Card 2: Single Day Query
        queryCardsPanel.add(createQueryCard(
            "📅 Single Day View",
            "View attendance for a specific day",
            ACCENT_GREEN,
            new Component[] {
                createDaySelector("Select Day", 1, 31),
                createStyledButton("📋 Show Day", ACCENT_GREEN, "queryDay")
            }
        ));
        
        // Card 3: Date Range Query
        queryCardsPanel.add(createQueryCard(
            "📆 Date Range Analysis",
            "Analyze attendance across multiple days",
            ACCENT_PURPLE,
            new Component[] {
                createRangeSelector("From", "To", 1, 31),
                createStyledButton("📈 Show Range", ACCENT_PURPLE, "queryRange")
            }
        ));
    }
    
    private JPanel createQueryCard(String title, String description, Color accentColor, Component[] components) {
        JPanel card = new JPanel();
        card.setLayout(new BorderLayout(0, 15));
        card.setBackground(CARD_BG);
        card.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(BORDER_COLOR, 1),
            new EmptyBorder(25, 25, 25, 25)
        ));
        
        // Add subtle shadow effect
        card.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createEmptyBorder(0, 0, 3, 3),
            BorderFactory.createCompoundBorder(
                new LineBorder(BORDER_COLOR, 1),
                new EmptyBorder(25, 25, 25, 25)
            )
        ));
        
        // Card header with accent border
        JPanel headerPanel = new JPanel(new BorderLayout());
        headerPanel.setBackground(CARD_BG);
        headerPanel.setBorder(BorderFactory.createCompoundBorder(
            new MatteBorder(0, 0, 3, 0, accentColor),
            new EmptyBorder(0, 0, 10, 0)
        ));
        
        JLabel titleLabel = new JLabel(title);
        titleLabel.setFont(new Font("Segoe UI", Font.BOLD, 16));
        titleLabel.setForeground(new Color(40, 50, 65));
        
        JLabel descLabel = new JLabel(description);
        descLabel.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        descLabel.setForeground(new Color(120, 130, 145));
        
        headerPanel.add(titleLabel, BorderLayout.NORTH);
        headerPanel.add(descLabel, BorderLayout.SOUTH);
        
        // Card content
        JPanel contentPanel = new JPanel();
        contentPanel.setLayout(new BoxLayout(contentPanel, BoxLayout.Y_AXIS));
        contentPanel.setBackground(CARD_BG);
        
        for (Component comp : components) {
            contentPanel.add(comp);
            contentPanel.add(Box.createVerticalStrut(12));
        }
        
        card.add(headerPanel, BorderLayout.NORTH);
        card.add(contentPanel, BorderLayout.CENTER);
        
        return card;
    }
    
    private JPanel createInputField(String label, String placeholder, int width) {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(CARD_BG);
        
        JLabel fieldLabel = new JLabel(label);
        fieldLabel.setFont(new Font("Segoe UI", Font.BOLD, 13));
        fieldLabel.setForeground(new Color(80, 90, 105));
        
        empCodeField = new JTextField(placeholder);
        empCodeField.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        empCodeField.setPreferredSize(new Dimension(width, 42));
        empCodeField.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(BORDER_COLOR, 1),
            new EmptyBorder(0, 15, 0, 15)
        ));
        empCodeField.setBackground(new Color(250, 251, 252));
        
        // Placeholder styling
        empCodeField.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                if (empCodeField.getText().equals(placeholder)) {
                    empCodeField.setText("");
                    empCodeField.setForeground(new Color(30, 40, 55));
                }
            }
            
            public void focusLost(java.awt.event.FocusEvent evt) {
                if (empCodeField.getText().isEmpty()) {
                    empCodeField.setText(placeholder);
                    empCodeField.setForeground(new Color(160, 170, 185));
                }
            }
        });
        
        panel.add(fieldLabel, BorderLayout.NORTH);
        panel.add(empCodeField, BorderLayout.CENTER);
        
        return panel;
    }
    
    private JPanel createDaySelector(String label, int min, int max) {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(CARD_BG);
        
        JLabel fieldLabel = new JLabel(label);
        fieldLabel.setFont(new Font("Segoe UI", Font.BOLD, 13));
        fieldLabel.setForeground(new Color(80, 90, 105));
        
        daySpinner = new JSpinner(new SpinnerNumberModel(1, min, max, 1));
        styleSpinner(daySpinner);
        
        panel.add(fieldLabel, BorderLayout.NORTH);
        panel.add(daySpinner, BorderLayout.CENTER);
        
        return panel;
    }
    
    private JPanel createRangeSelector(String fromLabel, String toLabel, int min, int max) {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(CARD_BG);
        
        JLabel fieldLabel = new JLabel("Date Range");
        fieldLabel.setFont(new Font("Segoe UI", Font.BOLD, 13));
        fieldLabel.setForeground(new Color(80, 90, 105));
        
        JPanel rangePanel = new JPanel(new GridLayout(2, 2, 10, 8));
        rangePanel.setBackground(CARD_BG);
        
        JLabel fromText = new JLabel(fromLabel);
        fromText.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        fromText.setForeground(new Color(100, 110, 125));
        
        startDaySpinner = new JSpinner(new SpinnerNumberModel(1, min, max, 1));
        styleSpinner(startDaySpinner);
        
        JLabel toText = new JLabel(toLabel);
        toText.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        toText.setForeground(new Color(100, 110, 125));
        
        endDaySpinner = new JSpinner(new SpinnerNumberModel(max, min, max, 1));
        styleSpinner(endDaySpinner);
        
        rangePanel.add(fromText);
        rangePanel.add(startDaySpinner);
        rangePanel.add(toText);
        rangePanel.add(endDaySpinner);
        
        panel.add(fieldLabel, BorderLayout.NORTH);
        panel.add(rangePanel, BorderLayout.CENTER);
        
        return panel;
    }
    
    private void styleSpinner(JSpinner spinner) {
        spinner.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        JSpinner.NumberEditor editor = new JSpinner.NumberEditor(spinner, "#");
        spinner.setEditor(editor);
        
        JTextField textField = editor.getTextField();
        textField.setPreferredSize(new Dimension(80, 42));
        textField.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(BORDER_COLOR, 1),
            new EmptyBorder(0, 15, 0, 15)
        ));
        textField.setBackground(new Color(250, 251, 252));
        textField.setHorizontalAlignment(JTextField.CENTER);
    }
    
    private JButton createStyledButton(String text, Color color, String action) {
        JButton button = new JButton(text);
        button.setFont(new Font("Segoe UI", Font.BOLD, 13));
        button.setBackground(color);
        button.setForeground(Color.WHITE);
        button.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(color.darker(), 1),
            new EmptyBorder(12, 25, 12, 25)
        ));
        button.setFocusPainted(false);
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        
        // Add action listener based on action type
        switch (action) {
            case "queryEmp":
                queryEmpBtn = button;
                button.setEnabled(false);
                button.addActionListener(this::showEmployeeData);
                break;
            case "queryDay":
                queryDayBtn = button;
                button.setEnabled(false);
                button.addActionListener(this::showDayData);
                break;
            case "queryRange":
                queryRangeBtn = button;
                button.setEnabled(false);
                button.addActionListener(this::showRangeData);
                break;
        }
        
        // Hover effect
        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                if (button.isEnabled()) {
                    button.setBackground(color.brighter());
                }
            }
            
            public void mouseExited(java.awt.event.MouseEvent evt) {
                if (button.isEnabled()) {
                    button.setBackground(color);
                }
            }
        });
        
        // Disabled state styling
        button.setEnabled(false);
        button.setBackground(new Color(220, 225, 230));
        button.setForeground(new Color(180, 185, 190));
        button.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(new Color(210, 215, 220), 1),
            new EmptyBorder(12, 25, 12, 25)
        ));
        
        return button;
    }
    
    private void setupResultsPanel() {
        resultsPanel = new JPanel(new BorderLayout(0, 15));
        resultsPanel.setBackground(LIGHT_BG);
        
        // Results header
        JPanel resultsHeader = new JPanel(new BorderLayout());
        resultsHeader.setBackground(LIGHT_BG);
        
        JLabel resultsTitle = new JLabel("📋 Query Results");
        resultsTitle.setFont(new Font("Segoe UI", Font.BOLD, 18));
        resultsTitle.setForeground(new Color(40, 50, 65));
        
        JPanel actionPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        actionPanel.setBackground(LIGHT_BG);
        
        clearBtn = createActionButton("🗑️ Clear Results", new Color(160, 170, 185));
        clearBtn.addActionListener(this::clearResults);
        
        exportBtn = createActionButton("📥 Export Data", ACCENT_PURPLE);
        exportBtn.setEnabled(false);
        
        actionPanel.add(clearBtn);
        actionPanel.add(exportBtn);
        
        resultsHeader.add(resultsTitle, BorderLayout.WEST);
        resultsHeader.add(actionPanel, BorderLayout.EAST);
        
        // Table setup
        setupTable();
        
        resultsPanel.add(resultsHeader, BorderLayout.NORTH);
        resultsPanel.add(tableScrollPane, BorderLayout.CENTER);
    }
    
    private JButton createActionButton(String text, Color color) {
        JButton button = new JButton(text);
        button.setFont(new Font("Segoe UI", Font.BOLD, 13));
        button.setBackground(color);
        button.setForeground(Color.WHITE);
        button.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(color.darker(), 1),
            new EmptyBorder(10, 20, 10, 20)
        ));
        button.setFocusPainted(false);
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        
        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                if (button.isEnabled()) {
                    button.setBackground(color.brighter());
                }
            }
            
            public void mouseExited(java.awt.event.MouseEvent evt) {
                if (button.isEnabled()) {
                    button.setBackground(color);
                }
            }
        });
        
        return button;
    }
    
    private void setupTable() {
        dataModel = new DefaultTableModel() {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };
        
        dataTable = new JTable(dataModel);
        dataTable.setToolTipText("Click column headers to sort • Double-click rows for details");
        
        // Table styling
        dataTable.setShowGrid(true);
        dataTable.setGridColor(new Color(240, 242, 245));
        dataTable.setRowHeight(40);
        dataTable.setAutoCreateRowSorter(true);
        dataTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        dataTable.setIntercellSpacing(new Dimension(0, 0));
        
        // Custom header
        JTableHeader header = dataTable.getTableHeader();
        header.setBackground(TABLE_HEADER_BG);
        header.setForeground(TABLE_HEADER_TEXT);
        header.setFont(new Font("Segoe UI", Font.BOLD, 13));
        header.setPreferredSize(new Dimension(header.getWidth(), 45));
        header.setReorderingAllowed(false);
        
        // Custom cell renderer
        dataTable.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            private static final long serialVersionUID = 1L;
            
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, 
                    boolean isSelected, boolean hasFocus, int row, int column) {
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                
                String text = displayValue(value);
                setText(text);
                setHorizontalAlignment(SwingConstants.CENTER);
                setBorder(new EmptyBorder(0, 12, 0, 12));
                setFont(new Font("Segoe UI", Font.PLAIN, 13));
                
                if (isSelected) {
                    c.setBackground(new Color(230, 240, 255));
                    setForeground(new Color(30, 40, 55));
                } else {
                    // Alternate row colors
                    if (row % 2 == 0) {
                        c.setBackground(Color.WHITE);
                    } else {
                        c.setBackground(new Color(250, 251, 252));
                    }
                    
                    // Status-based text coloring
                    try {
                        int statusCol = -1;
                        for (int i = 0; i < table.getColumnCount(); i++) {
                            if (table.getColumnName(i).equalsIgnoreCase("Status")) {
                                statusCol = i;
                                break;
                            }
                        }
                        
                        if (statusCol != -1) {
                            Object statusObj = table.getValueAt(row, statusCol);
                            String status = statusObj != null ? statusObj.toString().toUpperCase() : "";
                            
                            if (status.contains("A") || status.contains("ABSENT")) {
                                setForeground(new Color(220, 53, 69));
                                setFont(new Font("Segoe UI", Font.BOLD, 13));
                            } else if (status.contains("H") || status.contains("HALF")) {
                                setForeground(new Color(255, 153, 0));
                                setFont(new Font("Segoe UI", Font.BOLD, 13));
                            } else if (status.contains("P") || status.contains("PRESENT")) {
                                setForeground(new Color(40, 167, 69));
                                setFont(new Font("Segoe UI", Font.BOLD, 13));
                            } else {
                                setForeground(new Color(80, 90, 105));
                            }
                        }
                    } catch (Exception e) {
                        setForeground(new Color(80, 90, 105));
                    }
                }
                
                return c;
            }
        });
        
        tableScrollPane = new JScrollPane(dataTable);
        tableScrollPane.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(BORDER_COLOR, 1),
            BorderFactory.createEmptyBorder(1, 1, 1, 1)
        ));
        tableScrollPane.getViewport().setBackground(Color.WHITE);
        
        // Custom scrollbar
        JScrollBar verticalScrollBar = tableScrollPane.getVerticalScrollBar();
        verticalScrollBar.setUnitIncrement(16);
        verticalScrollBar.setBackground(Color.WHITE);
    }
    
    private JPanel createStatusBar() {
        JPanel statusBar = new JPanel(new BorderLayout());
        statusBar.setBackground(new Color(245, 247, 250));
        statusBar.setBorder(BorderFactory.createCompoundBorder(
            new MatteBorder(1, 0, 0, 0, BORDER_COLOR),
            new EmptyBorder(12, 20, 12, 20)
        ));
        
        statusLabel = new JLabel(" ");
        statusLabel.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        statusLabel.setForeground(new Color(100, 110, 125));
        
        JLabel helpLabel = new JLabel("💡 Tip: Process data in Configuration panel first");
        helpLabel.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        helpLabel.setForeground(new Color(140, 150, 165));
        
        statusBar.add(statusLabel, BorderLayout.WEST);
        statusBar.add(helpLabel, BorderLayout.EAST);
        
        return statusBar;
    }
    
    // ========== BUSINESS LOGIC METHODS ==========
    
    private void showEmployeeData(ActionEvent e) {
        if (mergeResult == null || queryViewer == null) {
            showError("📭 No data loaded", "Please process attendance data in the Configuration panel first.");
            return;
        }
        
        String empCode = empCodeField.getText().trim();
        if (empCode.isEmpty() || empCode.equals("Enter employee code...")) {
            showError("❓ Missing Employee ID", "Please enter a valid employee code.");
            return;
        }
        
        try {
            queryViewer.showByEmployee(empCode);
            updateStatus("✅ Showing attendance for employee: " + empCode);
        } catch (Exception ex) {
            showError("🚨 Query Error", "Failed to load employee data: " + ex.getMessage());
            ex.printStackTrace();
        }
    }
    
    private void showDayData(ActionEvent e) {
        if (mergeResult == null || queryViewer == null) {
            showError("📭 No data loaded", "Please process attendance data in the Configuration panel first.");
            return;
        }
        
        int day = (Integer) daySpinner.getValue();
        if (day < 1 || day > mergeResult.monthDays) {
            showError("⚠️ Invalid Day", "Day must be between 1 and " + mergeResult.monthDays);
            return;
        }
        
        try {
            queryViewer.showByDate(day);
            String monthName = LocalDate.of(currentYear, currentMonth, 1)
                .getMonth().getDisplayName(java.time.format.TextStyle.FULL, Locale.ENGLISH);
            updateStatus("📅 Showing attendance for " + monthName + " " + day + ", " + currentYear);
        } catch (Exception ex) {
            showError("🚨 Query Error", "Failed to load day data: " + ex.getMessage());
            ex.printStackTrace();
        }
    }
    
    private void showRangeData(ActionEvent e) {
        if (mergeResult == null || queryViewer == null) {
            showError("📭 No data loaded", "Please process attendance data in the Configuration panel first.");
            return;
        }
        
        int startDay = (Integer) startDaySpinner.getValue();
        int endDay = (Integer) endDaySpinner.getValue();
        
        if (startDay > endDay) {
            showError("⚠️ Invalid Range", "Start day must be less than or equal to end day.");
            return;
        }
        
        if (startDay < 1 || endDay > mergeResult.monthDays) {
            showError("⚠️ Out of Range", "Days must be between 1 and " + mergeResult.monthDays);
            return;
        }
        
        try {
            queryViewer.showByDateRange(startDay, endDay, currentYear, currentMonth);
            String monthName = LocalDate.of(currentYear, currentMonth, 1)
                .getMonth().getDisplayName(java.time.format.TextStyle.FULL, Locale.ENGLISH);
            updateStatus("📆 Showing attendance for " + monthName + " " + startDay + "-" + endDay + ", " + currentYear);
        } catch (Exception ex) {
            showError("🚨 Query Error", "Failed to load range data: " + ex.getMessage());
            ex.printStackTrace();
        }
    }
    
    private void clearResults(ActionEvent e) {
        dataModel.setRowCount(0);
        dataModel.setColumnCount(0);
        empCodeField.setText("Enter employee code...");
        empCodeField.setForeground(new Color(160, 170, 185));
        exportBtn.setEnabled(false);
        updateStatus("🔄 Results cleared. Ready for new queries.");
    }
    
    private void updateStatus(String message) {
        statusLabel.setText(" " + message);
    }
    
    private void showError(String title, String message) {
        JOptionPane.showMessageDialog(this, 
            "<html><div style='width: 300px; padding: 10px;'>" +
            "<h3 style='margin-top: 0; color: #dc3545;'>" + title + "</h3>" +
            "<p style='color: #666;'>" + message + "</p>" +
            "</div></html>",
            "Query Error", 
            JOptionPane.ERROR_MESSAGE);
        updateStatus("❌ " + title.toLowerCase());
    }
    
    private String displayValue(Object v) {
        if (v == null) return "-";
        String s = String.valueOf(v);
        if (s.trim().isEmpty()) return "-";
        return s;
    }
    
    // ========== PUBLIC METHODS ==========
    
    public void setMergeResult(AttendanceMerger.MergeResult mergeResult) {
        this.mergeResult = mergeResult;
        
        if (mergeResult != null) {
            this.currentYear = mergeResult.year;
            this.currentMonth = mergeResult.month;
            
            // Update spinners with correct month days
            SpinnerNumberModel dayModel = new SpinnerNumberModel(1, 1, mergeResult.monthDays, 1);
            SpinnerNumberModel startModel = new SpinnerNumberModel(1, 1, mergeResult.monthDays, 1);
            SpinnerNumberModel endModel = new SpinnerNumberModel(mergeResult.monthDays, 1, mergeResult.monthDays, 1);
            
            daySpinner.setModel(dayModel);
            startDaySpinner.setModel(startModel);
            endDaySpinner.setModel(endModel);
            
            // Enable query buttons
            enableQueryButtons(true);
            
            // Update stats panel
            updateStatsPanel(mergeResult);
            
            // Update panel title
            String monthName = LocalDate.of(currentYear, currentMonth, 1)
                .getMonth().getDisplayName(java.time.format.TextStyle.FULL, Locale.ENGLISH);
            panelTitleLabel.setText("Data Query Explorer - " + monthName + " " + currentYear);
            
            // Initialize query viewer
            initializeQueryViewer();
        }
    }
    
    private void enableQueryButtons(boolean enabled) {
        queryEmpBtn.setEnabled(enabled);
        queryDayBtn.setEnabled(enabled);
        queryRangeBtn.setEnabled(enabled);
        
        // Update button styles
        Color[] colors = {ACCENT_BLUE, ACCENT_GREEN, ACCENT_PURPLE};
        JButton[] buttons = {queryEmpBtn, queryDayBtn, queryRangeBtn};
        
        for (int i = 0; i < buttons.length; i++) {
            if (enabled) {
                buttons[i].setBackground(colors[i]);
                buttons[i].setForeground(Color.WHITE);
                buttons[i].setBorder(BorderFactory.createCompoundBorder(
                    new LineBorder(colors[i].darker(), 1),
                    new EmptyBorder(12, 25, 12, 25)
                ));
            } else {
                buttons[i].setBackground(new Color(220, 225, 230));
                buttons[i].setForeground(new Color(180, 185, 190));
                buttons[i].setBorder(BorderFactory.createCompoundBorder(
                    new LineBorder(new Color(210, 215, 220), 1),
                    new EmptyBorder(12, 25, 12, 25)
                ));
            }
        }
    }
    
    private void updateStatsPanel(AttendanceMerger.MergeResult result) {
        // Clear existing stats
        statsPanel.removeAll();
        
        // Add updated stats
        addStatItem("👥 Employees", String.valueOf(result.allEmployees.size()), ACCENT_BLUE);
        addStatItem("📅 Month Days", String.valueOf(result.monthDays), ACCENT_GREEN);
        
        String monthName = LocalDate.of(result.year, result.month, 1)
            .getMonth().getDisplayName(java.time.format.TextStyle.FULL, Locale.ENGLISH);
        addStatItem("🗓️ Period", monthName + " " + result.year, ACCENT_ORANGE);
        
        statsPanel.revalidate();
        statsPanel.repaint();
    }
    
    private void initializeQueryViewer() {
        try {
            String mergedPath = mergedFilePath != null ? mergedFilePath : "Merged_Attendance.xlsx";
            File mergedFile = new File(mergedPath);
            if (mergedFile.exists()) {
                queryViewer = new AttendanceQueryViewer(
                    mergedPath, 
                    mergeResult.allEmployees,
                    currentYear, 
                    currentMonth
                );
                updateStatus("✅ Ready to query " + mergeResult.allEmployees.size() + " employees");
            } else {
                updateStatus("⚠️ Merged file not found: " + mergedPath);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            updateStatus("❌ Error initializing query viewer: " + ex.getMessage());
        }
    }
    
    public void setMergedFilePath(String path) {
        this.mergedFilePath = path;
        if (path != null) {
            updateStatus("📁 Using merged file: " + (new File(path)).getName());
        }
    }
    
    public void enableExport(boolean enabled) {
        exportBtn.setEnabled(enabled);
        exportBtn.setBackground(enabled ? ACCENT_PURPLE : new Color(220, 225, 230));
        exportBtn.setForeground(enabled ? Color.WHITE : new Color(180, 185, 190));
    }
    
    public DefaultTableModel getDataModel() {
        return dataModel;
    }
    
    public JTable getDataTable() {
        return dataTable;
    }
    
    public boolean hasData() {
        return dataModel.getRowCount() > 0;
    }
}