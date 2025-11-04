package org.bioparse.cleaning;

import javax.swing.*;
import javax.swing.border.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.table.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*; // **FIXED: SINGLE import for AWT Color/Font**
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import com.formdev.flatlaf.FlatLightLaf;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.YearMonth;
import java.time.LocalDate;
import java.time.LocalTime;

public class AttendanceAppGUI extends JFrame {

	// ==================== SERIAL VERSION UID ====================
	private static final long serialVersionUID = 1L;

	// ==================== **FIXED: ALL COLOR CONSTANTS DEFINED**
	// ====================
	private static final java.awt.Color PRIMARY_COLOR = new java.awt.Color(41, 128, 185); // #2980B9
	private static final java.awt.Color SECONDARY_COLOR = new java.awt.Color(52, 152, 219); // #3498DB
	private static final java.awt.Color BACKGROUND_COLOR = new java.awt.Color(245, 245, 245); // #F5F5F5
	private static final java.awt.Color SIDEBAR_COLOR = new java.awt.Color(250, 250, 250);
	private static final java.awt.Color TEXT_PRIMARY = new java.awt.Color(51, 51, 51);
	private static final java.awt.Color BORDER_COLOR = new java.awt.Color(224, 224, 224);
	private static final java.awt.Color SUCCESS_COLOR = new java.awt.Color(39, 174, 96);
	private static final java.awt.Color WARNING_COLOR = new java.awt.Color(243, 156, 18);
	private static final java.awt.Color DANGER_COLOR = new java.awt.Color(231, 76, 60);

	// ==================== **FIXED: ALL FONT CONSTANTS DEFINED**
	// ====================
	private static final java.awt.Font FONT_BODY = new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 13);
	private static final java.awt.Font FONT_BUTTON = new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 13);
	private static final java.awt.Font FONT_SIDEBAR = new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 13);

	// ==================== FIELDS ====================
	private CardLayout cardLayout;
	private JPanel mainPanel, sidebarPanel;
	private JTextField inputFileField, mergedFileField, reportFileField;
	private JButton selectInputBtn, selectMergedBtn, selectReportBtn, processBtn, exportCsvBtn, exportMergedBtn;
	private JTextField empCodeField, fromDayField, toDayField;
	private JButton queryBtn;
	private JTable dataTable, reportTable;
	private DefaultTableModel dataModel, reportModel;
	private AttendanceMerger.MergeResult mergeResult;
	private AttendanceReportGenerator reportGenerator = new AttendanceReportGenerator();
	private AttendanceQueryViewer queryViewer;

	// **NEW: Export Filtered Data Button Field**
	private JButton exportFilteredBtn;
	private JPanel visualizationPanel;
	private JPanel contentPlaceholder;

	// Report filter UI
	private JPanel reportFilterPanel;
	private JTextField reportGlobalSearchField;
	private JComboBox<String> statusFilterCombo;
	private TableRowSorter<DefaultTableModel> reportSorter;
	private int statusColumnIndex = -1;

	// Dashboard stats
	private JLabel totalEmployeesLabel, presentTodayLabel, absentTodayLabel, lateArrivalsLabel;

	// NEW: Calendar components for holiday selection
	private JComboBox<Integer> yearComboBox;
	private JComboBox<String> monthComboBox;
	private JSpinner daySpinner;
	private JList<String> holidayList;
	private DefaultListModel<String> holidayListModel;
	private JButton addHolidayBtn, removeHolidayBtn;

	public AttendanceAppGUI() {
		try {
			UIManager.setLookAndFeel(new FlatLightLaf());
		} catch (Exception e) {
			e.printStackTrace();
		}

		setTitle("Attendance Analytics Dashboard");
		setSize(1400, 900);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLayout(new BorderLayout());
		getContentPane().setBackground(BACKGROUND_COLOR);

		// Enhanced Menu Bar
		setupMenuBar();

		// Enhanced Sidebar Navigation
		setupSidebar();

		// Main Content with CardLayout
		cardLayout = new CardLayout();
		mainPanel = new JPanel(cardLayout);
		mainPanel.setBackground(BACKGROUND_COLOR);

		mainPanel.add(createHomePanel(), "HOME");
		mainPanel.add(createConfigWizardPanel(), "CONFIG");
		mainPanel.add(createDataViewerPanel(), "DATA");
		mainPanel.add(createReportViewerPanel(), "REPORT");
		add(mainPanel, BorderLayout.CENTER);
		mainPanel.add(createVisualizationPanel(), "VISUALIZATION");

		showPanel("HOME");
	}

	private void setupMenuBar() {
		JMenuBar menuBar = new JMenuBar();
		menuBar.setBackground(java.awt.Color.WHITE);
		menuBar.setBorder(new EmptyBorder(5, 0, 5, 0));

		JMenu fileMenu = new JMenu("File");
		fileMenu.setFont(FONT_BODY);

		JMenuItem openItem = new JMenuItem("Open Input...");
		JMenuItem saveItem = new JMenuItem("Save Report...");
		openItem.setFont(FONT_BODY);
		saveItem.setFont(FONT_BODY);

		fileMenu.add(openItem);
		fileMenu.add(saveItem);

		JMenu helpMenu = new JMenu("Help");
		helpMenu.setFont(FONT_BODY);

		JMenuItem aboutItem = new JMenuItem("About");
		aboutItem.setFont(FONT_BODY);
		helpMenu.add(aboutItem);

		menuBar.add(fileMenu);
		menuBar.add(helpMenu);
		setJMenuBar(menuBar);
	}

	private void setupSidebar() {
		sidebarPanel = new JPanel();
		sidebarPanel.setLayout(new BoxLayout(sidebarPanel, BoxLayout.Y_AXIS));
		sidebarPanel.setBackground(SIDEBAR_COLOR);
		sidebarPanel.setBorder(
				new CompoundBorder(new MatteBorder(0, 0, 0, 1, BORDER_COLOR), new EmptyBorder(20, 10, 20, 10)));
		sidebarPanel.setPreferredSize(new Dimension(220, 0));

		// Enhanced sidebar header
		JLabel logoLabel = new JLabel("Attendance Pro");
		logoLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 20));
		logoLabel.setForeground(PRIMARY_COLOR);
		logoLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
		sidebarPanel.add(logoLabel);
		sidebarPanel.add(Box.createRigidArea(new Dimension(0, 30)));

		// Enhanced sidebar buttons with proper text wrapping
		addSidebarButton("Dashboard", "🏠", e -> showPanel("HOME"));
		addSidebarButton("Configuration", "⚙️", e -> showPanel("CONFIG"));
		addSidebarButton("Data Viewer", "📊", e -> showPanel("DATA"));
		addSidebarButton("Reports", "📈", e -> showPanel("REPORT"));
		addSidebarButton("Analytics & Charts", "📈", e -> showPanel("VISUALIZATION"));

		sidebarPanel.add(Box.createVerticalGlue());
		add(sidebarPanel, BorderLayout.WEST);
	}

	private void addSidebarButton(String text, String emoji, java.awt.event.ActionListener listener) {
		JButton button = new JButton("<html><div style='text-align: left;'>" + emoji + " " + text + "</div></html>");
		button.addActionListener(listener);
		button.setFont(FONT_SIDEBAR);
		button.setBackground(java.awt.Color.WHITE);
		button.setForeground(TEXT_PRIMARY);
		button.setBorder(
				new CompoundBorder(new MatteBorder(0, 0, 1, 0, BORDER_COLOR), new EmptyBorder(12, 10, 12, 10)));
		button.setHorizontalAlignment(SwingConstants.LEFT);
		button.setCursor(new Cursor(Cursor.HAND_CURSOR));
		button.setToolTipText("Go to " + text);

		// Hover effect
		button.addMouseListener(new java.awt.event.MouseAdapter() {
			public void mouseEntered(java.awt.event.MouseEvent evt) {
				button.setBackground(new java.awt.Color(240, 240, 240));
				button.setForeground(PRIMARY_COLOR);
			}

			public void mouseExited(java.awt.event.MouseEvent evt) {
				button.setBackground(java.awt.Color.WHITE);
				button.setForeground(TEXT_PRIMARY);
			}
		});

		button.setMaximumSize(new Dimension(200, 50));
		button.setPreferredSize(new Dimension(200, 50));
		sidebarPanel.add(button);
		sidebarPanel.add(Box.createRigidArea(new Dimension(0, 5)));
	}

	private void showPanel(String panelName) {
		cardLayout.show(mainPanel, panelName);
		updateDashboardStats(); // Update stats when showing dashboard
	}

	private JPanel createHomePanel() {
		JPanel panel = new JPanel(new BorderLayout());
		panel.setBackground(BACKGROUND_COLOR);
		panel.setBorder(new EmptyBorder(20, 20, 20, 20));

		// Header
		JLabel headerLabel = new JLabel("Attendance Analytics Dashboard");
		headerLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 28));
		headerLabel.setForeground(TEXT_PRIMARY);
		headerLabel.setBorder(new EmptyBorder(0, 0, 30, 0));
		panel.add(headerLabel, BorderLayout.NORTH);

		// Main content with image and stats
		JPanel contentPanel = new JPanel(new GridBagLayout());
		contentPanel.setBackground(BACKGROUND_COLOR);
		GridBagConstraints gbc = new GridBagConstraints();

		// Left side - Image
		gbc.gridx = 0;
		gbc.gridy = 0;
		gbc.weightx = 0.6;
		gbc.weighty = 1.0;
		gbc.fill = GridBagConstraints.BOTH;
		gbc.insets = new Insets(0, 0, 0, 20);

		JPanel imagePanel = createImagePanel();
		contentPanel.add(imagePanel, gbc);

		// Right side - Stats cards
		gbc.gridx = 1;
		gbc.gridy = 0;
		gbc.weightx = 0.4;
		gbc.insets = new Insets(0, 0, 0, 0);

		JPanel statsPanel = createStatsPanel();
		contentPanel.add(statsPanel, gbc);

		panel.add(contentPanel, BorderLayout.CENTER);
		return panel;
	}

	private JPanel createImagePanel() {
		JPanel panel = createCardPanel();
		panel.setLayout(new BorderLayout());

		try {
			// Load the image from resources
			ImageIcon originalIcon = new ImageIcon(getClass().getResource("/attendance-management-software-image.jpg"));
			if (originalIcon.getIconWidth() > 0) {
				// Scale the image to fit the panel
				Image scaledImage = originalIcon.getImage().getScaledInstance(600, 400, Image.SCALE_SMOOTH);
				ImageIcon scaledIcon = new ImageIcon(scaledImage);
				JLabel imageLabel = new JLabel(scaledIcon);
				imageLabel.setHorizontalAlignment(SwingConstants.CENTER);
				panel.add(imageLabel, BorderLayout.CENTER);
			} else {
				// Fallback if image not found
				JLabel fallbackLabel = new JLabel(
						"<html><div style='text-align: center; color: #666; font-size: 16px;'>"
								+ "🏢 Attendance Management System<br>"
								+ "<span style='font-size: 12px;'>Professional biometric attendance tracking solution</span>"
								+ "</div></html>",
						SwingConstants.CENTER);
				fallbackLabel.setFont(FONT_BODY);
				panel.add(fallbackLabel, BorderLayout.CENTER);
			}
		} catch (Exception e) {
			// Fallback if image loading fails
			JLabel fallbackLabel = new JLabel("<html><div style='text-align: center; color: #666; font-size: 16px;'>"
					+ "🏢 Attendance Management System<br>"
					+ "<span style='font-size: 12px;'>Professional biometric attendance tracking solution</span>"
					+ "</div></html>", SwingConstants.CENTER);
			fallbackLabel.setFont(FONT_BODY);
			panel.add(fallbackLabel, BorderLayout.CENTER);
		}

		return panel;
	}

	private JPanel createStatsPanel() {
		JPanel panel = new JPanel();
		panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
		panel.setBackground(BACKGROUND_COLOR);

		// Stats cards
		totalEmployeesLabel = createStatCard("Total Employees", "0", PRIMARY_COLOR, "👥");
		presentTodayLabel = createStatCard("Present Today", "0", SUCCESS_COLOR, "✅");
		absentTodayLabel = createStatCard("Absent Today", "0", WARNING_COLOR, "❌");
		lateArrivalsLabel = createStatCard("Late Arrivals", "0", DANGER_COLOR, "⏰");

		panel.add(totalEmployeesLabel);
		panel.add(Box.createRigidArea(new Dimension(0, 15)));
		panel.add(presentTodayLabel);
		panel.add(Box.createRigidArea(new Dimension(0, 15)));
		panel.add(absentTodayLabel);
		panel.add(Box.createRigidArea(new Dimension(0, 15)));
		panel.add(lateArrivalsLabel);

		return panel;
	}

	private JLabel createStatCard(String title, String value, java.awt.Color color, String emoji) {
		JLabel card = new JLabel(
				"<html><div style='background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; text-align: center;'>"
						+ "<div style='font-size: 24px; margin-bottom: 8px;'>" + emoji + "</div>"
						+ "<div style='font-size: 14px; color: #666; margin-bottom: 8px;'>" + title + "</div>"
						+ "<div style='font-size: 32px; font-weight: bold; color: "
						+ String.format("#%02x%02x%02x", color.getRed(), color.getGreen(), color.getBlue()) + ";'>"
						+ value + "</div>" + "</div></html>");
		card.setOpaque(false);
		return card;
	}

	private void updateDashboardStats() {
		// Update stats based on loaded data
		if (mergeResult != null) {
			int total = mergeResult.allEmployees.size();
			int present = (int) (total * 0.85); // Simulate 85% attendance
			int absent = total - present;
			int late = (int) (present * 0.15); // Simulate 15% late arrivals

			totalEmployeesLabel.setText(
					"<html><div style='background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; text-align: center;'>"
							+ "<div style='font-size: 24px; margin-bottom: 8px;'>👥</div>"
							+ "<div style='font-size: 14px; color: #666; margin-bottom: 8px;'>Total Employees</div>"
							+ "<div style='font-size: 32px; font-weight: bold; color: #2980B9;'>" + total + "</div>"
							+ "</div></html>");

			presentTodayLabel.setText(
					"<html><div style='background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; text-align: center;'>"
							+ "<div style='font-size: 24px; margin-bottom: 8px;'>✅</div>"
							+ "<div style='font-size: 14px; color: #666; margin-bottom: 8px;'>Present Today</div>"
							+ "<div style='font-size: 32px; font-weight: bold; color: #27AE60;'>" + present + "</div>"
							+ "</div></html>");

			absentTodayLabel.setText(
					"<html><div style='background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; text-align: center;'>"
							+ "<div style='font-size: 24px; margin-bottom: 8px;'>❌</div>"
							+ "<div style='font-size: 14px; color: #666; margin-bottom: 8px;'>Absent Today</div>"
							+ "<div style='font-size: 32px; font-weight: bold; color: #F39C12;'>" + absent + "</div>"
							+ "</div></html>");

			lateArrivalsLabel.setText(
					"<html><div style='background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; text-align: center;'>"
							+ "<div style='font-size: 24px; margin-bottom: 8px;'>⏰</div>"
							+ "<div style='font-size: 14px; color: #666; margin-bottom: 8px;'>Late Arrivals</div>"
							+ "<div style='font-size: 32px; font-weight: bold; color: #E74C3C;'>" + late + "</div>"
							+ "</div></html>");
		}
	}

	// ==================== ENHANCED DATA VIEWER WITH FILTERS ====================

	private JPanel createDataViewerPanel() {
		JPanel viewerPanel = new JPanel(new BorderLayout());
		viewerPanel.setBackground(BACKGROUND_COLOR);
		viewerPanel.setBorder(new EmptyBorder(10, 10, 10, 10));

		// Enhanced query panel with more filters
		JPanel queryPanel = createCardPanel();
		queryPanel.setLayout(new GridLayout(2, 1, 10, 10));

		// First row of filters
		JPanel filterRow1 = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
		filterRow1.setBackground(java.awt.Color.WHITE);

		filterRow1.add(createStyledLabel("Employee Code:"));
		empCodeField = createStyledTextField(10);
		filterRow1.add(empCodeField);

		filterRow1.add(createStyledLabel("From Day:"));
		fromDayField = createStyledTextField(3);
		filterRow1.add(fromDayField);

		filterRow1.add(createStyledLabel("To Day:"));
		toDayField = createStyledTextField(3);
		filterRow1.add(toDayField);

		// Second row of filters
		JPanel filterRow2 = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
		filterRow2.setBackground(java.awt.Color.WHITE);

		filterRow2.add(createStyledLabel("Status:"));
		JComboBox<String> statusFilter = createStyledComboBox();
		statusFilter
				.setModel(new DefaultComboBoxModel<>(new String[] { "All", "Present", "Absent", "Late", "Half Day" }));
		filterRow2.add(statusFilter);

		filterRow2.add(createStyledLabel("Department:"));
		JComboBox<String> deptFilter = createStyledComboBox();
		deptFilter.setModel(new DefaultComboBoxModel<>(new String[] { "All", "IT", "HR", "Finance", "Operations" }));
		filterRow2.add(deptFilter);

		queryBtn = createPrimaryButton("🔍 Query Data");
		queryBtn.addActionListener(this::queryData);
		queryBtn.setEnabled(false);
		filterRow2.add(queryBtn);

		JButton clearFiltersBtn = createSecondaryButton("Clear Filters");
		clearFiltersBtn.addActionListener(e -> clearDataFilters());
		filterRow2.add(clearFiltersBtn);

		// **FIXED: EXPORT FILTERED DATA BUTTON**
		exportFilteredBtn = createSuccessButton("📊 Export Filtered Data");
		exportFilteredBtn.addActionListener(e -> exportFilteredData());
		exportFilteredBtn.setEnabled(false);
		filterRow2.add(exportFilteredBtn);

		queryPanel.add(filterRow1);
		queryPanel.add(filterRow2);
		viewerPanel.add(queryPanel, BorderLayout.NORTH);

		dataModel = new DefaultTableModel();
		dataTable = new JTable(dataModel);
		dataTable.setToolTipText("Attendance data table - Click column headers to sort");

		configureTableAppearance(dataTable);

		// Add table sorter for data table
		dataTable.setAutoCreateRowSorter(true);

		viewerPanel.add(new JScrollPane(dataTable), BorderLayout.CENTER);

		return viewerPanel;
	}

	/** **UPDATED: Clear Data Filters Method** */
	private void clearDataFilters() {
		empCodeField.setText("");
		fromDayField.setText("");
		toDayField.setText("");

		// Clear table and disable export button
		dataModel.setRowCount(0);
		if (exportFilteredBtn != null) {
			exportFilteredBtn.setEnabled(false);
		}

		if (dataTable.getRowSorter() != null) {
			TableRowSorter<? extends TableModel> sorter = (TableRowSorter<? extends TableModel>) dataTable
					.getRowSorter();
			sorter.setRowFilter(null);
		}
	}

	/** **UPDATED: Query Data Method - ENABLES EXPORT BUTTON** */
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
			// **FIX: Enable export button when data is loaded**
			exportFilteredBtn.setEnabled(true);
			return;
		} else if (from == to) {
			queryViewer.showByDate(from);
			// **FIX: Enable export button when data is loaded**
			exportFilteredBtn.setEnabled(true);
			return;
		} else {
			Object[] columns = { "Employee Code", "Employee Name", "Day", "InTime", "OutTime", "Duration", "Status" };
			dataModel.setColumnIdentifiers(columns);

			int rowCount = 0; // Track rows added
			for (AttendanceMerger.EmployeeData emp : mergeResult.allEmployees) {
				for (int d = from; d <= to; d++) {
					String in = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d - 1));
					String out = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("OutTime"), d - 1));
					String dur = calculateDuration(in, out);
					String stat = displayValue(AttendanceUtils.safeGet(emp.dailyData.get("Status"), d - 1));

					dataModel.addRow(new Object[] { emp.empId, emp.empName, d, in, out, dur, stat });
					rowCount++;
				}
			}

			TableColumnModel cm = dataTable.getColumnModel();
			for (int i = 0; i < cm.getColumnCount(); i++) {
				TableColumn col = cm.getColumn(i);
				DefaultTableCellRenderer renderer = new DefaultTableCellRenderer();
				if (i == 1)
					renderer.setHorizontalAlignment(SwingConstants.LEFT);
				else
					renderer.setHorizontalAlignment(SwingConstants.CENTER);
				col.setCellRenderer(renderer);
			}

			// **CRITICAL FIX: Enable export button when data is loaded**
			if (rowCount > 0) {
				exportFilteredBtn.setEnabled(true);
			} else {
				exportFilteredBtn.setEnabled(false);
				JOptionPane.showMessageDialog(this, "No data found for the selected criteria!", "No Data",
						JOptionPane.WARNING_MESSAGE);
			}
		}
	}

	private String calculateDuration(String inTime, String outTime) {
		if (inTime == null || inTime.isEmpty() || outTime == null || outTime.isEmpty() || "-".equals(inTime)
				|| "-".equals(outTime)) {
			return "-";
		}
		try {
			// Parse times (assuming format like "09:30", "17:45")
			LocalTime in = LocalTime.parse(inTime);
			LocalTime out = LocalTime.parse(outTime);

			long minutes = java.time.Duration.between(in, out).toMinutes();
			if (minutes < 0) {
				minutes += 24 * 60; // Handle next-day outTime
			}
			long hours = minutes / 60;
			long mins = minutes % 60;
			return String.format("%d:%02d", hours, mins);
		} catch (Exception e) {
			return "-"; // Return dash if parsing fails
		}
	}

	/** **UPDATED: Export Filtered Data Method** */
	private void exportFilteredData() {
		// Check if there's any visible (filtered) data
		int visibleRowCount = dataTable.getRowCount();
		if (visibleRowCount == 0) {
			JOptionPane.showMessageDialog(this, "No filtered data to export!", "Export Error",
					JOptionPane.WARNING_MESSAGE);
			return;
		}

		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setDialogTitle("Export Filtered Data to Excel");
		fileChooser.setSelectedFile(new File("filtered_attendance_data_"
				+ LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("yyyyMMdd")) + ".xlsx"));

		int userSelection = fileChooser.showSaveDialog(this);

		if (userSelection == JFileChooser.APPROVE_OPTION) {
			File fileToSave = fileChooser.getSelectedFile();

			// Ensure .xlsx extension
			if (!fileToSave.getName().toLowerCase().endsWith(".xlsx")) {
				fileToSave = new File(fileToSave.getAbsolutePath() + ".xlsx");
			}

			try {
				exportFilteredDataToExcel(fileToSave);
				JOptionPane
						.showMessageDialog(this,
								String.format("✅ %d rows of filtered data exported successfully!\n%s", visibleRowCount,
										fileToSave.getAbsolutePath()),
								"Export Success", JOptionPane.INFORMATION_MESSAGE);
			} catch (Exception ex) {
				JOptionPane.showMessageDialog(this, "❌ Error exporting filtered data: " + ex.getMessage(),
						"Export Error", JOptionPane.ERROR_MESSAGE);
				ex.printStackTrace();
			}
		}
	}

	/**
	 * **FIXED: Actual implementation of exporting filtered data to Excel**
	 */
	private void exportFilteredDataToExcel(File outputFile) throws Exception {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Filtered Attendance Data");

		TableModel model = dataTable.getModel();
		int rowCount = model.getRowCount();
		int colCount = model.getColumnCount();

		// Create header row with style **FIXED: CORRECT POI FONT USAGE**
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// **FIXED: Create POI Font correctly**
		org.apache.poi.ss.usermodel.Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerStyle.setFont(headerFont);

		headerStyle.setBorderTop(BorderStyle.THIN);
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setBorderLeft(BorderStyle.THIN);
		headerStyle.setBorderRight(BorderStyle.THIN);

		Row headerRow = sheet.createRow(0);
		for (int col = 0; col < colCount; col++) {
			Cell cell = headerRow.createCell(col);
			cell.setCellValue(model.getColumnName(col));
			cell.setCellStyle(headerStyle);
		}

		// Create data rows with styling
		CellStyle dataStyle = workbook.createCellStyle();
		dataStyle.setBorderTop(BorderStyle.THIN);
		dataStyle.setBorderBottom(BorderStyle.THIN);
		dataStyle.setBorderLeft(BorderStyle.THIN);
		dataStyle.setBorderRight(BorderStyle.THIN);
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

		// Auto-size columns for better readability
		for (int col = 0; col < colCount; col++) {
			sheet.autoSizeColumn(col);
		}

		// Write to file
		try (FileOutputStream fos = new FileOutputStream(outputFile)) {
			workbook.write(fos);
		}

		workbook.close();
	}

	// ==================== ENHANCED REPORT VIEWER ====================

	private JPanel createReportViewerPanel() {
		JPanel reportViewerPanel = new JPanel(new BorderLayout());
		reportViewerPanel.setBackground(BACKGROUND_COLOR);
		reportViewerPanel.setBorder(new EmptyBorder(10, 10, 10, 10));

		reportModel = new DefaultTableModel();
		reportTable = new JTable(reportModel);
		reportTable.setToolTipText("Employee metrics report - Click column headers to sort");

		configureTableAppearance(reportTable);

		// Enhanced filter panel with more options
		reportFilterPanel = createCardPanel();
		reportFilterPanel.setLayout(new GridLayout(2, 1, 10, 10));

		JPanel searchRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
		searchRow.setBackground(java.awt.Color.WHITE);
		searchRow.add(createStyledLabel("🔍 Search:"));
		reportGlobalSearchField = createStyledTextField(25);
		reportGlobalSearchField.setToolTipText("Search across all columns");
		searchRow.add(reportGlobalSearchField);

		JPanel filterRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
		filterRow.setBackground(java.awt.Color.WHITE);
		filterRow.add(createStyledLabel("Status:"));
		statusFilterCombo = createStyledComboBox();
		statusFilterCombo.setModel(new DefaultComboBoxModel<>(new String[] { "All" }));
		filterRow.add(statusFilterCombo);

		filterRow.add(createStyledLabel("Department:"));
		JComboBox<String> reportDeptFilter = createStyledComboBox();
		reportDeptFilter
				.setModel(new DefaultComboBoxModel<>(new String[] { "All", "IT", "HR", "Finance", "Operations" }));
		filterRow.add(reportDeptFilter);

		JButton exportPdfBtn = createSuccessButton("📄 Export PDF");
		exportPdfBtn.addActionListener(e -> exportToPdf());
		filterRow.add(exportPdfBtn);

		reportFilterPanel.add(searchRow);
		reportFilterPanel.add(filterRow);
		reportFilterPanel.setVisible(false);
		reportViewerPanel.add(reportFilterPanel, BorderLayout.NORTH);

		reportViewerPanel.add(new JScrollPane(reportTable), BorderLayout.CENTER);
		return reportViewerPanel;
	}

	private void debugVisualizationData() {
		System.out.println("=== VISUALIZATION DEBUG INFO ===");
		System.out.println("MergeResult: " + (mergeResult != null ? "Present" : "Null"));
		if (mergeResult != null) {
			System.out.println("Employees count: " + mergeResult.allEmployees.size());
			System.out.println("Month Days: " + mergeResult.monthDays);
			System.out.println("Year: " + mergeResult.year);
			System.out.println("Month: " + mergeResult.month);

			if (!mergeResult.allEmployees.isEmpty()) {
				AttendanceMerger.EmployeeData sampleEmp = mergeResult.allEmployees.get(0);
				System.out.println("Sample Employee: " + sampleEmp.empId + " - " + sampleEmp.empName);
				System.out.println("Data keys: " + sampleEmp.dailyData.keySet());

				// Check if we have status data
				List<String> statusList = sampleEmp.dailyData.get("Status");
				if (statusList != null) {
					System.out.println("Status data count: " + statusList.size());
					System.out.println(
							"First 10 status values: " + statusList.subList(0, Math.min(10, statusList.size())));
				} else {
					System.out.println("❌ STATUS DATA IS NULL!");
				}

				// Check other data types
				for (String key : Arrays.asList("InTime", "OutTime", "Duration")) {
					List<String> dataList = sampleEmp.dailyData.get(key);
					if (dataList != null) {
						System.out.println(key + " data count: " + dataList.size());
					} else {
						System.out.println("❌ " + key + " DATA IS NULL!");
					}
				}
			} else {
				System.out.println("❌ EMPLOYEE LIST IS EMPTY!");
			}
		} else {
			System.out.println("❌ MERGE RESULT IS NULL - Data not processed yet!");
		}
		System.out.println("=================================");
	}

	private void exportToPdf() {
		JOptionPane.showMessageDialog(this, "PDF export feature would be implemented here!");
	}

	// ==================== **FIXED: STYLING UTILITIES** ====================

	private JPanel createCardPanel() {
		JPanel card = new JPanel();
		card.setBackground(java.awt.Color.WHITE);
		card.setBorder(new CompoundBorder(new LineBorder(BORDER_COLOR, 1, true), new EmptyBorder(15, 15, 15, 15)));
		return card;
	}

	private JLabel createStyledLabel(String text) {
		JLabel label = new JLabel(text);
		label.setFont(FONT_BODY);
		label.setForeground(TEXT_PRIMARY);
		return label;
	}

	private JTextField createStyledTextField(int columns) {
		JTextField field = new JTextField(columns);
		field.setFont(FONT_BODY);
		field.setBorder(
				BorderFactory.createCompoundBorder(new LineBorder(BORDER_COLOR, 1), new EmptyBorder(8, 10, 8, 10)));
		return field;
	}

	private JComboBox<String> createStyledComboBox() {
		JComboBox<String> combo = new JComboBox<>();
		combo.setFont(FONT_BODY);
		combo.setBackground(java.awt.Color.WHITE);
		combo.setBorder(
				BorderFactory.createCompoundBorder(new LineBorder(BORDER_COLOR, 1), new EmptyBorder(5, 10, 5, 10)));
		return combo;
	}

	private JButton createPrimaryButton(String text) {
		JButton button = new JButton(text);
		button.setFont(FONT_BUTTON);
		button.setBackground(PRIMARY_COLOR);
		button.setForeground(java.awt.Color.WHITE);
		button.setBorder(BorderFactory.createCompoundBorder(new LineBorder(PRIMARY_COLOR, 1, true),
				new EmptyBorder(10, 15, 10, 15)));
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

	private JButton createSecondaryButton(String text) {
		JButton button = new JButton(text);
		button.setFont(FONT_BUTTON);
		button.setBackground(java.awt.Color.WHITE);
		button.setForeground(TEXT_PRIMARY);
		button.setBorder(BorderFactory.createCompoundBorder(new LineBorder(BORDER_COLOR, 1, true),
				new EmptyBorder(8, 12, 8, 12)));
		button.setCursor(new Cursor(Cursor.HAND_CURSOR));
		button.setFocusPainted(false);

		button.addMouseListener(new java.awt.event.MouseAdapter() {
			public void mouseEntered(java.awt.event.MouseEvent evt) {
				button.setBackground(new java.awt.Color(245, 245, 245));
			}

			public void mouseExited(java.awt.event.MouseEvent evt) {
				button.setBackground(java.awt.Color.WHITE);
			}
		});

		return button;
	}

	private JButton createSuccessButton(String text) {
		JButton button = new JButton(text);
		button.setFont(FONT_BUTTON);
		button.setBackground(SUCCESS_COLOR);
		button.setForeground(java.awt.Color.WHITE);
		button.setBorder(BorderFactory.createCompoundBorder(new LineBorder(SUCCESS_COLOR, 1, true),
				new EmptyBorder(10, 15, 10, 15)));
		button.setCursor(new Cursor(Cursor.HAND_CURSOR));
		button.setFocusPainted(false);

		button.addMouseListener(new java.awt.event.MouseAdapter() {
			public void mouseEntered(java.awt.event.MouseEvent evt) {
				button.setBackground(new java.awt.Color(46, 204, 113));
			}

			public void mouseExited(java.awt.event.MouseEvent evt) {
				button.setBackground(SUCCESS_COLOR);
			}
		});

		return button;
	}

	// Create the visualization panel method
private JPanel createVisualizationPanel() {
    visualizationPanel = new JPanel(new BorderLayout());
    visualizationPanel.setBackground(BACKGROUND_COLOR);
    visualizationPanel.setBorder(new EmptyBorder(10, 10, 10, 10));

    // Header
    JPanel headerPanel = new JPanel(new BorderLayout());
    JLabel headerLabel = new JLabel("Analytics & Visualization Dashboard");
    headerLabel.setFont(new Font("Segoe UI", Font.BOLD, 28));
    headerLabel.setForeground(TEXT_PRIMARY);
    
    JButton debugBtn = new JButton("Debug");
    debugBtn.addActionListener(e -> {
        debugVisualizationData();
        rebuildVisualization(); // <-- Force rebuild
    });
    
    headerPanel.add(headerLabel, BorderLayout.CENTER);
    headerPanel.add(debugBtn, BorderLayout.EAST);
    headerPanel.setBorder(new EmptyBorder(0, 0, 20, 0));
    visualizationPanel.add(headerPanel, BorderLayout.NORTH);

    // Placeholder content
    contentPlaceholder = new JPanel(new BorderLayout());
    contentPlaceholder.add(createNoDataPanel("Process attendance data to see charts."), BorderLayout.CENTER);
    visualizationPanel.add(contentPlaceholder, BorderLayout.CENTER);

    return visualizationPanel;
}
private JPanel createDashboardTab() {
    JPanel container = new JPanel(new BorderLayout());
    container.setBackground(BACKGROUND_COLOR);

    if (mergeResult != null && !mergeResult.allEmployees.isEmpty()) {
        JPanel gridPanel = new JPanel(new GridLayout(3, 3, 15, 15));
        gridPanel.setBackground(BACKGROUND_COLOR);

        // Add all 9 charts
        gridPanel.add(createChartCard("Trend", ChartDataGenerator.createAttendanceTrendChart(mergeResult.allEmployees, mergeResult.monthDays, mergeResult.year, mergeResult.month)));
        gridPanel.add(createChartCard("Distribution", ChartDataGenerator.createAttendanceDistributionChart(mergeResult.allEmployees, mergeResult.monthDays)));
        gridPanel.add(createChartCard("Department", ChartDataGenerator.createDepartmentWiseAttendanceChart(mergeResult.allEmployees)));
        gridPanel.add(createChartCard("Performance", ChartDataGenerator.createTopPerformersChart(mergeResult.allEmployees, mergeResult.monthDays)));
        gridPanel.add(createChartCard("Late Arrivals", ChartDataGenerator.createLateArrivalsAnalysis(mergeResult.allEmployees, mergeResult.monthDays)));
        gridPanel.add(createChartCard("Overtime", ChartDataGenerator.createOvertimeAnalysisChart(mergeResult.allEmployees)));
        gridPanel.add(createChartCard("Daily Trend", ChartDataGenerator.createDailyAttendanceLineChart(mergeResult.allEmployees, mergeResult.monthDays)));
        gridPanel.add(createChartCard("Weekly Pattern", ChartDataGenerator.createWeeklyAttendanceChart(mergeResult.allEmployees, mergeResult.monthDays)));
        gridPanel.add(createChartCard("Shift Compliance", ChartDataGenerator.createShiftComplianceChart(mergeResult.allEmployees)));

        // Wrap in JScrollPane
        JScrollPane scrollPane = new JScrollPane(gridPanel);
        scrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        scrollPane.getVerticalScrollBar().setUnitIncrement(16); // Smooth scroll

        container.add(scrollPane, BorderLayout.CENTER);
    } else {
        container.add(createNoDataPanel("Process data first to see charts."), BorderLayout.CENTER);
    }

    return container;
}

private void rebuildVisualization() {
    if (contentPlaceholder == null) return;

    contentPlaceholder.removeAll();

    if (mergeResult == null || mergeResult.allEmployees.isEmpty()) {
        contentPlaceholder.add(createNoDataPanel("Process attendance data first to see charts and analytics"), BorderLayout.CENTER);
    } else {
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setFont(FONT_BODY);

        tabbedPane.addTab("Dashboard", createDashboardTab());
        tabbedPane.addTab("Department Analysis", createDepartmentAnalysisTab());
        tabbedPane.addTab("Employee Performance", createEmployeePerformanceTab());
        tabbedPane.addTab("Overtime", createOvertimeAnalysisTab());

        contentPlaceholder.add(tabbedPane, BorderLayout.CENTER);
    }

    contentPlaceholder.revalidate();
    contentPlaceholder.repaint();
}

private JPanel createDepartmentAnalysisTab() {
		JPanel panel = new JPanel(new BorderLayout());
		panel.setBackground(BACKGROUND_COLOR);

		if (mergeResult != null) {
			JPanel deptChart = ChartDataGenerator.createDepartmentWiseAttendanceChart(mergeResult.allEmployees);
			deptChart.setPreferredSize(new Dimension(800, 500));
			panel.add(deptChart, BorderLayout.CENTER);
		} else {
			panel.add(createNoDataPanel(), BorderLayout.CENTER);
		}

		return panel;
	}

	private JPanel createEmployeePerformanceTab() {
		JPanel panel = new JPanel(new GridLayout(1, 2, 15, 15));
		panel.setBackground(BACKGROUND_COLOR);
		panel.setBorder(new EmptyBorder(10, 10, 10, 10));

		if (mergeResult != null) {
			// Top Performers
			JPanel performersChart = ChartDataGenerator.createTopPerformersChart(mergeResult.allEmployees,
					mergeResult.monthDays);
			performersChart.setPreferredSize(new Dimension(400, 500));
			panel.add(createChartCard("Top Performers", performersChart));

			// Late Arrivals
			JPanel lateChart = ChartDataGenerator.createLateArrivalsAnalysis(mergeResult.allEmployees,
					mergeResult.monthDays);
			lateChart.setPreferredSize(new Dimension(400, 500));
			panel.add(createChartCard("Late Arrivals", lateChart));
		} else {
			panel.add(createNoDataPanel());
		}

		return panel;
	}

//	private JPanel createTrendAnalysisTab() {
//		JPanel panel = new JPanel(new BorderLayout());
//		panel.setBackground(BACKGROUND_COLOR);
//
//		if (mergeResult != null) {
//			JPanel trendChart = ChartDataGenerator.createMonthlyTrendChart(mergeResult.allEmployees,
//					mergeResult.monthDays, mergeResult.year, mergeResult.month);
//			trendChart.setPreferredSize(new Dimension(800, 500));
//			panel.add(trendChart, BorderLayout.CENTER);
//		} else {
//			panel.add(createNoDataPanel(), BorderLayout.CENTER);
//		}
//
//		return panel;
//	}

	private JPanel createOvertimeAnalysisTab() {
		JPanel panel = new JPanel(new BorderLayout());
		panel.setBackground(BACKGROUND_COLOR);

		if (mergeResult != null) {
			JPanel overtimeChart = ChartDataGenerator.createOvertimeAnalysisChart(mergeResult.allEmployees);
			overtimeChart.setPreferredSize(new Dimension(800, 500));
			panel.add(overtimeChart, BorderLayout.CENTER);
		} else {
			panel.add(createNoDataPanel(), BorderLayout.CENTER);
		}

		return panel;
	}

	private JPanel createChartCard(String title, JPanel chartPanel) {
    JPanel card = createCardPanel();
    card.setLayout(new BorderLayout());

    JLabel titleLabel = new JLabel(title);
    titleLabel.setFont(FONT_BUTTON);
    titleLabel.setForeground(PRIMARY_COLOR);
    titleLabel.setBorder(new EmptyBorder(10, 10, 5, 10));
    card.add(titleLabel, BorderLayout.NORTH);

    // Make chart panel fill available space
    chartPanel.setPreferredSize(null); // Remove fixed size
    card.add(chartPanel, BorderLayout.CENTER);

    return card;
}
	private JPanel createNoDataPanel() {
		JPanel panel = new JPanel(new BorderLayout());
		panel.setBackground(BACKGROUND_COLOR);

		JLabel label = new JLabel("<html><div style='text-align: center; color: #666; font-size: 16px;'>"
				+ "📊<br>No data available for visualization<br>"
				+ "<span style='font-size: 12px;'>Process attendance data first to see charts and analytics</span>"
				+ "</div></html>", SwingConstants.CENTER);
		label.setFont(FONT_BODY);

		panel.add(label, BorderLayout.CENTER);
		return panel;
	}

	private JPanel createErrorPanel(String error) {
		JPanel panel = new JPanel(new BorderLayout());
		panel.setBackground(BACKGROUND_COLOR);

		JLabel label = new JLabel(
				"<html><div style='text-align: center; color: #d00; font-size: 14px;'>" + "❌<br>Chart Error<br>"
						+ "<span style='font-size: 12px;'>" + error + "</span>" + "</div></html>",
				SwingConstants.CENTER);
		label.setFont(FONT_BODY);

		panel.add(label, BorderLayout.CENTER);
		return panel;
	}

	private JPanel createNoDataPanel(String message) {
		JPanel panel = new JPanel(new BorderLayout());
		panel.setBackground(BACKGROUND_COLOR);

		JLabel label = new JLabel("<html><div style='text-align: center; color: #666; font-size: 16px;'>" + "📊<br>"
				+ message + "</div></html>", SwingConstants.CENTER);
		label.setFont(FONT_BODY);

		panel.add(label, BorderLayout.CENTER);
		return panel;
	}

	// ==================== ORIGINAL BUSINESS LOGIC METHODS ====================

	private JPanel createConfigWizardPanel() {
		JPanel wizardPanel = new JPanel(new CardLayout());
		CardLayout wizardLayout = (CardLayout) wizardPanel.getLayout();

		// Step 1: File Selection
		JPanel step1 = new JPanel(new GridBagLayout());
		GridBagConstraints gbc = new GridBagConstraints();
		gbc.insets = new Insets(10, 10, 10, 10);
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.weightx = 1.0;

		gbc.gridx = 0;
		gbc.gridy = 0;
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

		// Export Merged File Button
		gbc.gridy++;
		step1.add(new JLabel("Export Merged File:"), gbc);
		gbc.gridy++;
		exportMergedBtn = new JButton("Export Merged File");
		exportMergedBtn.addActionListener(this::exportMergedFile);
		exportMergedBtn.setEnabled(false);
		step1.add(exportMergedBtn, gbc);

		gbc.gridy++;
		JButton nextBtn = new JButton("Next");
		nextBtn.addActionListener(e -> wizardLayout.next(wizardPanel));
		step1.add(nextBtn, gbc);

		wizardPanel.add(step1, "STEP1");

		// Step 2: ENHANCED Holidays/Year/Month with Calendar
		JPanel step2 = createEnhancedConfigPanel(wizardLayout, wizardPanel);

		wizardPanel.add(step2, "STEP2");

		return wizardPanel;
	}

	private JPanel createEnhancedConfigPanel(CardLayout wizardLayout, JPanel wizardPanel) {
		JPanel step2 = new JPanel(new BorderLayout(10, 10));
		step2.setBorder(new EmptyBorder(20, 20, 20, 20));
		step2.setBackground(BACKGROUND_COLOR);

		// Header
		JLabel headerLabel = new JLabel("Configuration Settings");
		headerLabel.setFont(FONT_BUTTON);
		headerLabel.setForeground(PRIMARY_COLOR);
		step2.add(headerLabel, BorderLayout.NORTH);

		// Main content panel
		JPanel contentPanel = new JPanel(new GridLayout(1, 2, 20, 0));
		contentPanel.setBackground(BACKGROUND_COLOR);

		// Left panel - Date Selection
		JPanel dateSelectionPanel = createCardPanel();
		dateSelectionPanel.setLayout(new GridBagLayout());
		GridBagConstraints gbc = new GridBagConstraints();
		gbc.insets = new Insets(10, 10, 10, 10);
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.weightx = 1.0;

		// Year Selection
		gbc.gridx = 0;
		gbc.gridy = 0;
		dateSelectionPanel.add(createStyledLabel("Year:"), gbc);
		gbc.gridy++;
		yearComboBox = createYearComboBox();
		dateSelectionPanel.add(yearComboBox, gbc);

		// Month Selection
		gbc.gridy++;
		dateSelectionPanel.add(createStyledLabel("Month:"), gbc);
		gbc.gridy++;
		monthComboBox = createMonthComboBox();
		dateSelectionPanel.add(monthComboBox, gbc);

		// Day Selection
		gbc.gridy++;
		dateSelectionPanel.add(createStyledLabel("Day:"), gbc);
		gbc.gridy++;
		daySpinner = createDaySpinner();
		dateSelectionPanel.add(daySpinner, gbc);

		// Add Holiday Button
		gbc.gridy++;
		addHolidayBtn = createPrimaryButton("➕ Add Holiday");
		addHolidayBtn.addActionListener(this::addHoliday);
		dateSelectionPanel.add(addHolidayBtn, gbc);

		contentPanel.add(dateSelectionPanel);

		// Right panel - Holidays List
		JPanel holidaysPanel = createCardPanel();
		holidaysPanel.setLayout(new BorderLayout(10, 10));

		JLabel holidaysLabel = new JLabel("Selected Holidays:");
		holidaysLabel.setFont(FONT_BUTTON);
		holidaysLabel.setForeground(PRIMARY_COLOR);
		holidaysPanel.add(holidaysLabel, BorderLayout.NORTH);

		holidayListModel = new DefaultListModel<>();
		holidayList = new JList<>(holidayListModel);
		holidayList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
		holidayList.setFont(FONT_BODY);
		holidaysPanel.add(new JScrollPane(holidayList), BorderLayout.CENTER);

		// Remove button
		removeHolidayBtn = createSecondaryButton("➖ Remove Selected");
		removeHolidayBtn.addActionListener(e -> removeHoliday());
		holidaysPanel.add(removeHolidayBtn, BorderLayout.SOUTH);

		contentPanel.add(holidaysPanel);
		step2.add(contentPanel, BorderLayout.CENTER);

		// Bottom button panel
		JPanel buttonPanel = new JPanel(new FlowLayout());
		buttonPanel.setBackground(BACKGROUND_COLOR);

		JButton backBtn = createSecondaryButton("← Back");
		backBtn.addActionListener(e -> wizardLayout.previous(wizardPanel));

		processBtn = createPrimaryButton("🚀 Process Data");
		processBtn.addActionListener(this::processData);

		exportCsvBtn = createSuccessButton("📊 Export Report to CSV");
		exportCsvBtn.setEnabled(false);
		exportCsvBtn.addActionListener(this::exportCsv);

		buttonPanel.add(backBtn);
		buttonPanel.add(processBtn);
		buttonPanel.add(exportCsvBtn);

		step2.add(buttonPanel, BorderLayout.SOUTH);

		// Add listener to update day spinner when month/year changes
		monthComboBox.addActionListener(e -> updateDaySpinner());
		yearComboBox.addActionListener(e -> updateDaySpinner());

		return step2;
	}

	// Calendar component creation methods
	private JComboBox<Integer> createYearComboBox() {
		JComboBox<Integer> combo = new JComboBox<>();
		combo.setFont(FONT_BODY);
		combo.setBackground(java.awt.Color.WHITE);
		combo.setEditable(true);

		int currentYear = LocalDate.now().getYear();
		for (int year = 2000; year <= 2050; year++) {
			combo.addItem(year);
		}
		combo.setSelectedItem(currentYear);

		return combo;
	}

	private JComboBox<String> createMonthComboBox() {
		JComboBox<String> combo = new JComboBox<>();
		combo.setFont(FONT_BODY);
		combo.setBackground(java.awt.Color.WHITE);

		String[] months = { "January", "February", "March", "April", "May", "June", "July", "August", "September",
				"October", "November", "December" };
		for (String month : months) {
			combo.addItem(month);
		}
		combo.setSelectedIndex(LocalDate.now().getMonthValue() - 1);

		return combo;
	}

	private JSpinner createDaySpinner() {
		int currentYear = (Integer) yearComboBox.getSelectedItem();
		int currentMonth = monthComboBox.getSelectedIndex() + 1;
		int daysInMonth = YearMonth.of(currentYear, currentMonth).lengthOfMonth();

		SpinnerNumberModel model = new SpinnerNumberModel(1, 1, daysInMonth, 1);
		JSpinner spinner = new JSpinner(model);
		spinner.setFont(FONT_BODY);

		JSpinner.NumberEditor editor = new JSpinner.NumberEditor(spinner, "#");
		spinner.setEditor(editor);

		return spinner;
	}

	private void updateDaySpinner() {
		try {
			int year = (Integer) yearComboBox.getSelectedItem();
			int month = monthComboBox.getSelectedIndex() + 1;
			int daysInMonth = YearMonth.of(year, month).lengthOfMonth();

			int currentDay = (Integer) daySpinner.getValue();
			if (currentDay > daysInMonth) {
				currentDay = daysInMonth;
			}

			SpinnerNumberModel newModel = new SpinnerNumberModel(currentDay, 1, daysInMonth, 1);
			daySpinner.setModel(newModel);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void addHoliday(ActionEvent e) {
		try {
			int year = (Integer) yearComboBox.getSelectedItem();
			int month = monthComboBox.getSelectedIndex() + 1;
			int day = (Integer) daySpinner.getValue();

			if (!isValidDate(year, month, day)) {
				JOptionPane.showMessageDialog(this,
						"Invalid date: " + getMonthName(month) + " " + day + ", " + year + "\n" + getMonthName(month)
								+ " has only " + YearMonth.of(year, month).lengthOfMonth() + " days.",
						"Invalid Date", JOptionPane.ERROR_MESSAGE);
				return;
			}

			String holidayString = String.format("%s %d, %d", getMonthName(month), day, year);

			if (!holidayListModel.contains(holidayString)) {
				holidayListModel.addElement(holidayString);
			} else {
				JOptionPane.showMessageDialog(this, "This date is already in the holiday list.", "Duplicate Holiday",
						JOptionPane.WARNING_MESSAGE);
			}
		} catch (Exception ex) {
			JOptionPane.showMessageDialog(this, "Error adding holiday: " + ex.getMessage(), "Error",
					JOptionPane.ERROR_MESSAGE);
		}
	}

	private void removeHoliday() {
		int selectedIndex = holidayList.getSelectedIndex();
		if (selectedIndex != -1) {
			holidayListModel.remove(selectedIndex);
		} else {
			JOptionPane.showMessageDialog(this, "Please select a holiday to remove.", "No Selection",
					JOptionPane.WARNING_MESSAGE);
		}
	}

	private boolean isValidDate(int year, int month, int day) {
		try {
			YearMonth yearMonth = YearMonth.of(year, month);
			return day >= 1 && day <= yearMonth.lengthOfMonth();
		} catch (Exception e) {
			return false;
		}
	}

	private String getMonthName(int month) {
		String[] months = { "January", "February", "March", "April", "May", "June", "July", "August", "September",
				"October", "November", "December" };
		return months[month - 1];
	}

	private int getMonthNumber(String monthName) {
		String[] months = { "January", "February", "March", "April", "May", "June", "July", "August", "September",
				"October", "November", "December" };
		for (int i = 0; i < months.length; i++) {
			if (months[i].equalsIgnoreCase(monthName)) {
				return i + 1;
			}
		}
		return 1;
	}

	private void exportMergedFile(ActionEvent e) {
		if (mergeResult == null) {
			JOptionPane.showMessageDialog(this, "No merged data available. Process data first.");
			return;
		}

		JFileChooser chooser = new JFileChooser();
		chooser.setSelectedFile(new File("merged_attendance.xlsx"));
		int res = chooser.showSaveDialog(this);
		if (res == JFileChooser.APPROVE_OPTION) {
			try {
				String sourceFile = mergedFileField.getText();
				String destFile = chooser.getSelectedFile().getAbsolutePath();

				Files.copy(Paths.get(sourceFile), Paths.get(destFile), StandardCopyOption.REPLACE_EXISTING);
				JOptionPane.showMessageDialog(this, "Merged file exported successfully!");
			} catch (Exception ex) {
				JOptionPane.showMessageDialog(this, "Error exporting merged file: " + ex.getMessage(), "Error",
						JOptionPane.ERROR_MESSAGE);
			}
		}
	}

	private void configureTableAppearance(JTable table) {
		table.setShowGrid(true);
		table.setGridColor(new java.awt.Color(220, 220, 220));
		table.setIntercellSpacing(new Dimension(1, 1));
		table.setRowHeight(22);
		table.setAutoCreateRowSorter(true);

		DefaultTableCellRenderer defaultRenderer = new DefaultTableCellRenderer() {
			private static final long serialVersionUID = 1L;

			@Override
			public void setValue(Object value) {
				String text = displayValue(value);
				setText(text);
			}
		};
		defaultRenderer.setHorizontalAlignment(SwingConstants.CENTER);
		table.setDefaultRenderer(Object.class, defaultRenderer);
	}

	private String displayValue(Object v) {
		if (v == null)
			return "-";
		String s = String.valueOf(v);
		if (s.trim().isEmpty())
			return "-";
		return s;
	}

	private void selectFile(ActionEvent e) {
		JFileChooser chooser = new JFileChooser();
		int res = chooser.showOpenDialog(this);
		if (res == JFileChooser.APPROVE_OPTION) {
			File file = chooser.getSelectedFile();
			if (e.getSource() == selectInputBtn)
				inputFileField.setText(file.getAbsolutePath());
			else if (e.getSource() == selectMergedBtn)
				mergedFileField.setText(file.getAbsolutePath());
			else if (e.getSource() == selectReportBtn)
				reportFileField.setText(file.getAbsolutePath());
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

        // Extract holidays from the list
        List<Integer> holidays = new ArrayList<>();
        int selectedYear;
        Object yearObj = yearComboBox.getSelectedItem();
        if (yearObj instanceof Integer) {
            selectedYear = (Integer) yearObj;
        } else {
            try {
                selectedYear = Integer.parseInt(yearObj.toString());
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(this, "Invalid year format. Please enter a valid year.", 
                    "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
        }
        int selectedMonth = getMonthNumber((String) monthComboBox.getSelectedItem());
        
        for (int i = 0; i < holidayListModel.size(); i++) {
            String holidayStr = holidayListModel.getElementAt(i);
            String[] parts = holidayStr.split(" ");
            if (parts.length >= 2) {
                try {
                    String dayStr = parts[1].replace(",", "");
                    int day = Integer.parseInt(dayStr);
                    holidays.add(day);
                } catch (NumberFormatException ex) {
                    // Skip invalid entries
                }
            }
        }

        if (holidays.isEmpty()) {
            int result = JOptionPane.showConfirmDialog(this, 
                "No holidays selected. Continue without holidays?", 
                "No Holidays", JOptionPane.YES_NO_OPTION);
            if (result != JOptionPane.YES_OPTION) {
                return;
            }
        }

        // PROCESS THE DATA
        System.out.println("=== STARTING DATA PROCESSING ===");
        mergeResult = AttendanceMerger.merge(inputFile, mergedFile, holidays);
        mergeResult.year = selectedYear;
        mergeResult.month = selectedMonth;
        
        System.out.println("Processing complete:");
        System.out.println("- Employees: " + (mergeResult.allEmployees != null ? mergeResult.allEmployees.size() : "null"));
        System.out.println("- Month Days: " + mergeResult.monthDays);
        
        // Generate report
        reportGenerator.generate(mergedFile, reportFile, mergeResult.allEmployees, 
            mergeResult.monthDays, holidays, mergeResult.year, mergeResult.month);

        loadReportIntoTable(reportFile);

        // ENABLE ALL EXPORT BUTTONS AFTER PROCESSING
        exportCsvBtn.setEnabled(true);
        if (queryBtn != null) queryBtn.setEnabled(true);
        if (exportFilteredBtn != null) exportFilteredBtn.setEnabled(true);
        exportMergedBtn.setEnabled(true);

        if (queryViewer == null) {
            // Use the constructor that takes explicit year/month
            queryViewer = new AttendanceQueryViewer(mergedFile, mergeResult.allEmployees, selectedYear, selectedMonth);
        } else {
            // If queryViewer already exists, update it with new data
            queryViewer = new AttendanceQueryViewer(mergedFile, mergeResult.allEmployees, selectedYear, selectedMonth);
        }

        JOptionPane.showMessageDialog(this, 
            "Processing complete!\n" +
            "• Year: " + selectedYear + "\n" +
            "• Month: " + getMonthName(selectedMonth) + "\n" +
            "• Holidays: " + holidays.size() + " days\n" +
            "• Employees: " + mergeResult.allEmployees.size(), 
            "Success", JOptionPane.INFORMATION_MESSAGE);
        
     // ADD THIS:
        rebuildVisualization();  // <--- THIS IS THE KEY
            
        // Force refresh visualization
        refreshVisualizationData();
        updateDashboardStats();
        
    } catch (Exception ex) {
        ex.printStackTrace();
        JOptionPane.showMessageDialog(this, "Error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
    }
}
	private void exportCsv(ActionEvent e) {
		if (mergeResult == null)
			return;
		JFileChooser chooser = new JFileChooser();
		chooser.setSelectedFile(new File("report.csv"));
		int res = chooser.showSaveDialog(this);
		if (res == JFileChooser.APPROVE_OPTION) {
			try {
				reportGenerator.exportToCsv(chooser.getSelectedFile().getAbsolutePath(), mergeResult.allEmployees,
						mergeResult.monthDays, parseHolidays(), mergeResult.year, mergeResult.month);
				JOptionPane.showMessageDialog(this, "CSV exported successfully!");
			} catch (Exception ex) {
				JOptionPane.showMessageDialog(this, "Error exporting CSV: " + ex.getMessage(), "Error",
						JOptionPane.ERROR_MESSAGE);
			}
		}
	}

	// Add this method to refresh visualization data
	private void refreshVisualizationData() {
		if (visualizationPanel != null) {
			CardLayout layout = (CardLayout) mainPanel.getLayout();
			layout.show(mainPanel, "VISUALIZATION");
			// Force UI update
			visualizationPanel.revalidate();
			visualizationPanel.repaint();
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
			Cell cell = header.getCell(i);
			columns[i] = (cell == null) ? ("Col" + i) : cell.getStringCellValue();
		}
		reportModel.setColumnIdentifiers(columns);

		for (int r = 1; r <= sheet.getLastRowNum(); r++) {
			Row row = sheet.getRow(r);
			if (row == null)
				continue;
			Object[] data = new Object[row.getLastCellNum()];
			for (int c = 0; c < data.length; c++) {
				Cell cell = row.getCell(c);
				if (cell == null)
					data[c] = "-";
				else if (cell.getCellType() == CellType.NUMERIC) {
					double dv = cell.getNumericCellValue();
					if (dv == Math.rint(dv))
						data[c] = String.valueOf((long) dv);
					else
						data[c] = String.valueOf(dv);
				} else if (cell.getCellType() == CellType.STRING) {
					String s = cell.getStringCellValue();
					data[c] = (s == null || s.trim().isEmpty()) ? "-" : s;
				} else {
					String s = cell.toString();
					data[c] = (s == null || s.trim().isEmpty()) ? "-" : s;
				}
			}
			reportModel.addRow(data);
		}

		wb.close();
		fis.close();

		reportSorter = new TableRowSorter<>(reportModel);
		reportTable.setRowSorter(reportSorter);

		reportFilterPanel.setVisible(true);
		statusColumnIndex = -1;
		for (int i = 0; i < reportModel.getColumnCount(); i++) {
			if (String.valueOf(reportModel.getColumnName(i)).equalsIgnoreCase("Status")) {
				statusColumnIndex = i;
				break;
			}
		}

		if (statusColumnIndex >= 0) {
			Set<String> statuses = new LinkedHashSet<>();
			statuses.add("All");
			for (int r = 0; r < reportModel.getRowCount(); r++) {
				Object val = reportModel.getValueAt(r, statusColumnIndex);
				statuses.add(displayValue(val));
			}
			statusFilterCombo.setModel(new DefaultComboBoxModel<>(statuses.toArray(new String[0])));
			statusFilterCombo.setEnabled(true);
		} else {
			statusFilterCombo.setModel(new DefaultComboBoxModel<>(new String[] { "All" }));
			statusFilterCombo.setEnabled(false);
		}

		installReportFilters();
	}

	private void installReportFilters() {
		reportGlobalSearchField.getDocument().removeDocumentListener(globalSearchListener);
		reportGlobalSearchField.getDocument().addDocumentListener(globalSearchListener);

		statusFilterCombo.removeActionListener(statusComboListener);
		statusFilterCombo.addActionListener(statusComboListener);

		applyCombinedReportFilter();
	}

	private final DocumentListener globalSearchListener = new DocumentListener() {
		@Override
		public void insertUpdate(DocumentEvent e) {
			applyCombinedReportFilter();
		}

		@Override
		public void removeUpdate(DocumentEvent e) {
			applyCombinedReportFilter();
		}

		@Override
		public void changedUpdate(DocumentEvent e) {
			applyCombinedReportFilter();
		}
	};

	private final java.awt.event.ActionListener statusComboListener = e -> applyCombinedReportFilter();

	private void applyCombinedReportFilter() {
		if (reportSorter == null)
			return;

		List<RowFilter<Object, Object>> filters = new ArrayList<>();

		String text = reportGlobalSearchField.getText();
		if (text != null && !text.trim().isEmpty()) {
			String expr = "(?i).*" + Pattern.quote(text.trim()) + ".*";
			RowFilter<Object, Object> regexFilter = RowFilter.regexFilter(expr);
			filters.add(regexFilter);
		}

		if (statusColumnIndex >= 0 && statusFilterCombo.getSelectedItem() != null) {
			String sel = String.valueOf(statusFilterCombo.getSelectedItem());
			if (!"All".equalsIgnoreCase(sel)) {
				RowFilter<Object, Object> statusFilter = new RowFilter<Object, Object>() {
					private static final long serialVersionUID = 1L;

					@Override
					public boolean include(Entry<? extends Object, ? extends Object> entry) {
						Object v = entry.getValue(statusColumnIndex);
						String s = (v == null) ? "-" : String.valueOf(v);
						return s.equalsIgnoreCase(sel);
					}
				};
				filters.add(statusFilter);
			}
		}

		if (filters.isEmpty()) {
			reportSorter.setRowFilter(null);
		} else if (filters.size() == 1) {
			reportSorter.setRowFilter(filters.get(0));
		} else {
			reportSorter.setRowFilter(RowFilter.andFilter(filters));
		}
	}

	private List<Integer> parseHolidays() {
		List<Integer> holidays = new ArrayList<>();
		for (int i = 0; i < holidayListModel.size(); i++) {
			String holidayStr = holidayListModel.getElementAt(i);
			String[] parts = holidayStr.split(" ");
			if (parts.length >= 2) {
				try {
					String dayStr = parts[1].replace(",", "");
					holidays.add(Integer.parseInt(dayStr));
				} catch (NumberFormatException e) {
				}
			}
		}
		return holidays;
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> new AttendanceAppGUI().setVisible(true));
	}
}