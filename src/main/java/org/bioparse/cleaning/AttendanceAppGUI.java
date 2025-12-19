package org.bioparse.cleaning;

import javax.swing.*;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Locale;
import java.util.ArrayList;
import java.util.List;
import javax.swing.table.DefaultTableModel;

import com.formdev.flatlaf.FlatLightLaf;
import org.bioparse.cleaning.ui.*;

public class AttendanceAppGUI extends JFrame {

	private static final long serialVersionUID = 1L;

	private CardLayout cardLayout;
	private JPanel mainPanel;
	private DashboardPanel dashboardPanel;
	private ConfigWizardPanel configWizardPanel;
	private DataViewerPanel dataViewerPanel;
	private ReportViewerPanel reportViewerPanel;
	private VisualizationPanel visualizationPanel;

	// Shared data
	private AttendanceMerger.MergeResult mergeResult;

	public AttendanceAppGUI() {
		setupLookAndFeel();
		setupFrame();
		setupMenuBar();
		setupMainPanel();
		setupSidebar();

		setVisible(true);
	}

	private void setupLookAndFeel() {
		try {
			UIManager.setLookAndFeel(new FlatLightLaf());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void setupFrame() {
		setTitle("Attendance Analytics Dashboard");
		setSize(1400, 900);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLayout(new BorderLayout());
		getContentPane().setBackground(Constants.BACKGROUND_COLOR);
		setLocationRelativeTo(null); // Center the window
	}

	private void setupMenuBar() {
		JMenuBar menuBar = new JMenuBar();
		menuBar.setBackground(Color.WHITE);
		menuBar.setBorder(BorderFactory.createEmptyBorder(5, 0, 5, 0));

		JMenu fileMenu = new JMenu("File");
		fileMenu.setFont(Constants.FONT_BODY);

		JMenuItem openItem = new JMenuItem("Open Input...");
		JMenuItem saveItem = new JMenuItem("Save Report...");
		openItem.setFont(Constants.FONT_BODY);
		saveItem.setFont(Constants.FONT_BODY);

		fileMenu.add(openItem);
		fileMenu.add(saveItem);

		JMenu helpMenu = new JMenu("Help");
		helpMenu.setFont(Constants.FONT_BODY);

		JMenuItem aboutItem = new JMenuItem("About");
		aboutItem.setFont(Constants.FONT_BODY);
		helpMenu.add(aboutItem);

		menuBar.add(fileMenu);
		menuBar.add(helpMenu);
		setJMenuBar(menuBar);
	}

	private void setupMainPanel() {
		cardLayout = new CardLayout();
		mainPanel = new JPanel(cardLayout);
		mainPanel.setBackground(Constants.BACKGROUND_COLOR);

		// Initialize all panels
		dashboardPanel = new DashboardPanel();
		configWizardPanel = new ConfigWizardPanel(this::processData, this::exportCsv, this::exportMergedFile);
		dataViewerPanel = new DataViewerPanel(this::queryData, this::exportFilteredData);
		reportViewerPanel = new ReportViewerPanel();
		visualizationPanel = new VisualizationPanel();

		// Add panels to CardLayout with unique names
		mainPanel.add(dashboardPanel, "HOME");
		mainPanel.add(configWizardPanel, "CONFIG");
		mainPanel.add(dataViewerPanel, "DATA");
		mainPanel.add(reportViewerPanel, "REPORT");
		mainPanel.add(visualizationPanel, "VISUALIZATION");

		add(mainPanel, BorderLayout.CENTER);
	}

	private void setupSidebar() {
		SidebarPanel sidebarPanel = new SidebarPanel(e -> showPanel("HOME"), e -> showPanel("CONFIG"),
				e -> showPanel("DATA"), e -> showPanel("REPORT"), e -> showPanel("VISUALIZATION"));
		add(sidebarPanel, BorderLayout.WEST);
	}

	private void showPanel(String panelName) {
	    System.out.println("Showing panel: " + panelName);
	    cardLayout.show(mainPanel, panelName);
	    
	    // Update panels with latest data when shown
	    if ("HOME".equals(panelName) && dashboardPanel != null) {
	        dashboardPanel.updateStats(mergeResult);
	    } else if ("DATA".equals(panelName) && dataViewerPanel != null && mergeResult != null) {
	        dataViewerPanel.setMergeResult(mergeResult);
	        
	        // Also set merged file path if available
	        String mergedPath = configWizardPanel.getMergedFilePath();
	        if (mergedPath == null || mergedPath.isEmpty()) {
	            mergedPath = "Merged_Attendance.xlsx";
	        }
	        dataViewerPanel.setMergedFilePath(mergedPath);
	    } else if ("VISUALIZATION".equals(panelName) && visualizationPanel != null) {
	        visualizationPanel.setMergeResult(mergeResult);
	    } else if ("REPORT".equals(panelName) && reportViewerPanel != null) {
	        // Always generate report data directly
	        if (mergeResult != null && !reportViewerPanel.hasData()) {
	            generateAndLoadMonthlyReport();
	        } else if (!reportViewerPanel.hasData()) {
	            // Show empty state
	            reportViewerPanel.clearData();
	        }
	    }
	}

	private void processData(ActionEvent e) {
	    // Validate input file
	    String inputFile = configWizardPanel.getInputFilePath();
	    if (inputFile == null || inputFile.isEmpty() || !new File(inputFile).exists()) {
	        configWizardPanel.setStatusMessage("Please select a valid input file first!", true);
	        return;
	    }
	    
	    // Get output paths and make them effectively final
	    final String finalMergedFile;
	    String tempMergedFile = configWizardPanel.getMergedFilePath();
	    if (tempMergedFile == null || tempMergedFile.isEmpty()) {
	        finalMergedFile = "Merged_Attendance.xlsx";
	        configWizardPanel.setStatusMessage("Using default merged file name", false);
	    } else {
	        finalMergedFile = tempMergedFile;
	    }
	    
	    final String finalReportFile;
	    String tempReportFile = configWizardPanel.getReportFilePath();
	    if (tempReportFile == null || tempReportFile.isEmpty()) {
	        finalReportFile = "Monthly_Attendance_Report.xlsx";
	        configWizardPanel.setStatusMessage("Using default report file name", false);
	    } else {
	        finalReportFile = tempReportFile;
	    }
	    
	    // Get holidays
	    final List<Integer> finalHolidays = configWizardPanel.getHolidays();
	    
	    // Get selected month/year
	    final int finalYear = configWizardPanel.getSelectedYear();
	    final int finalMonth = configWizardPanel.getSelectedMonth();
	    
	    // Show loading indicator
	    configWizardPanel.showLoading();
	    configWizardPanel.setStatusMessage("Starting data processing...", false);
	    
	    // Run processing in a background thread
	    SwingWorker<Boolean, Integer> worker = new SwingWorker<Boolean, Integer>() {
	        @Override
	        protected Boolean doInBackground() throws Exception {
	            try {
	                publish(10); // 10% progress
	                
	                System.out.println("Starting merge process...");
	                System.out.println("Input: " + inputFile);
	                System.out.println("Output: " + finalMergedFile);
	                System.out.println("Year: " + finalYear + ", Month: " + finalMonth);
	                System.out.println("Holidays: " + finalHolidays);
	                
	                publish(30); // 30% progress
	                
	                // Call the AttendanceMerger
	                mergeResult = AttendanceMerger.merge(
	                    inputFile, 
	                    finalMergedFile, 
	                    finalHolidays
	                );
	                
	                publish(90); // 90% progress
	                
	                return true;
	            } catch (Exception ex) {
	                ex.printStackTrace();
	                return false;
	            }
	        }
	        
	        @Override
	        protected void process(List<Integer> chunks) {
	            // Update progress
	            for (Integer progress : chunks) {
	                configWizardPanel.showProgress(progress, "Processing");
	            }
	        }
	        
	        @Override
	        protected void done() {
	            try {
	                boolean success = get();
	                if (success) {
	                    // Update UI on success
	                    configWizardPanel.enableExportButtons(true);
	                    configWizardPanel.hideLoading();
	                    
	                    String monthName = LocalDate.of(finalYear, finalMonth, 1)
	                        .getMonth()
	                        .getDisplayName(TextStyle.FULL, Locale.ENGLISH);
	                    
	                    configWizardPanel.setStatusMessage(
	                        String.format("✓ Success! Processed %d employees for %s %d", 
	                            mergeResult.allEmployees.size(),
	                            monthName,
	                            finalYear), 
	                        false
	                    );
	                    
	                    // AUTO-GENERATE MONTHLY REPORT when processing completes
	                    try {
	                        // Generate monthly summary report file
	                        generateMonthlyReport(finalMergedFile, finalReportFile, finalHolidays, finalYear, finalMonth);
	                        
	                        System.out.println("Auto-generated monthly report: " + finalReportFile);
	                        
	                    } catch (Exception reportEx) {
	                        System.err.println("Failed to auto-generate report file: " + reportEx.getMessage());
	                        reportEx.printStackTrace();
	                    }
	                    
	                    // Generate and load data directly into ReportViewerPanel
	                    SwingUtilities.invokeLater(() -> {
	                        generateAndLoadMonthlyReport();
	                    });
	                    
	                    // Update dashboard panel
	                    if (dashboardPanel != null) {
	                        dashboardPanel.updateStats(mergeResult);
	                    }
	                    
	                    // Update data viewer panel with both merge result and file path
	                    if (dataViewerPanel != null) {
	                        dataViewerPanel.setMergeResult(mergeResult);
	                        
	                        // Get merged file path from config panel
	                        String mergedPath = configWizardPanel.getMergedFilePath();
	                        if (mergedPath == null || mergedPath.isEmpty()) {
	                            mergedPath = "Merged_Attendance.xlsx";
	                        }
	                        dataViewerPanel.setMergedFilePath(mergedPath);
	                        
	                        System.out.println("DataViewerPanel updated with merge result and file: " + mergedPath);
	                    }
	                    
	                    // Update visualization panel
	                    if (visualizationPanel != null) {
	                        visualizationPanel.setMergeResult(mergeResult);
	                    }
	                    
	                    // Log success
	                    System.out.println("Processing completed successfully!");
	                    System.out.println("Total employees: " + mergeResult.allEmployees.size());
	                    System.out.println("Month: " + finalMonth + "/" + finalYear);
	                    System.out.println("Days in month: " + mergeResult.monthDays);
	                    
	                    // Show success message
	                    JOptionPane.showMessageDialog(AttendanceAppGUI.this,
	                        String.format("Processing Complete!\n\n" +
	                            "✓ Employees: %d\n" +
	                            "✓ Month: %s %d\n" +
	                            "✓ Days: %d\n" +
	                            "✓ Merged file created\n" +
	                            "✓ Monthly report generated\n\n" +
	                            "You can now view the monthly summary in the Report panel.",
	                            mergeResult.allEmployees.size(),
	                            monthName,
	                            finalYear,
	                            mergeResult.monthDays),
	                        "Processing Complete",
	                        JOptionPane.INFORMATION_MESSAGE);
	                    
	                } else {
	                    configWizardPanel.hideLoading();
	                    configWizardPanel.setStatusMessage(
	                        "✗ Processing failed. Check console for errors.", 
	                        true
	                    );
	                    
	                    // Show error dialog
	                    JOptionPane.showMessageDialog(AttendanceAppGUI.this,
	                        "Processing failed. Please check the console for detailed errors.",
	                        "Processing Error",
	                        JOptionPane.ERROR_MESSAGE);
	                }
	            } catch (java.util.concurrent.ExecutionException ex) {
	                // Handle execution exception (actual error from doInBackground)
	                configWizardPanel.hideLoading();
	                String errorMessage = ex.getCause() != null ? ex.getCause().getMessage() : ex.getMessage();
	                configWizardPanel.setStatusMessage(
	                    "✗ Error: " + errorMessage, 
	                    true
	                );
	                
	                // Show detailed error dialog
	                String detailedError = "Processing error:\n" + errorMessage;
	                if (errorMessage.contains("File") || errorMessage.contains("path")) {
	                    detailedError += "\n\nPlease check if the input file exists and is accessible.";
	                }
	                
	                JOptionPane.showMessageDialog(AttendanceAppGUI.this,
	                    detailedError,
	                    "Processing Error",
	                    JOptionPane.ERROR_MESSAGE);
	                    
	                ex.printStackTrace();
	                
	            } catch (Exception ex) {
	                ex.printStackTrace();
	                configWizardPanel.hideLoading();
	                configWizardPanel.setStatusMessage(
	                    "✗ Error: " + ex.getMessage(), 
	                    true
	                );
	                
	                JOptionPane.showMessageDialog(AttendanceAppGUI.this,
	                    "Unexpected error: " + ex.getMessage(),
	                    "Error",
	                    JOptionPane.ERROR_MESSAGE);
	            }
	        }
	    };
	    
	    worker.execute();
	}

	private void exportCsv(ActionEvent e) {
	    if (mergeResult == null) {
	        configWizardPanel.setStatusMessage("Please process data first!", true);
	        return;
	    }
	    
	    // Get report file path from config panel
	    String reportFilePath = configWizardPanel.getReportFilePath();
	    if (reportFilePath == null || reportFilePath.isEmpty()) {
	        reportFilePath = "Monthly_Attendance_Report.xlsx";
	    }
	    
	    // Make sure it ends with .xlsx
	    if (!reportFilePath.toLowerCase().endsWith(".xlsx")) {
	        if (reportFilePath.toLowerCase().endsWith(".csv")) {
	            reportFilePath = reportFilePath.substring(0, reportFilePath.length() - 4) + ".xlsx";
	        } else {
	            reportFilePath = reportFilePath + ".xlsx";
	        }
	    }
	    
	    // Create final variables for use in inner class
	    final String finalReportFilePath = reportFilePath;
	    final String finalMergedFilePath;
	    
	    String tempMergedPath = configWizardPanel.getMergedFilePath();
	    if (tempMergedPath == null || tempMergedPath.isEmpty()) {
	        finalMergedFilePath = "Merged_Attendance.xlsx";
	    } else {
	        finalMergedFilePath = tempMergedPath;
	    }
	    
	    final List<Integer> finalHolidays = configWizardPanel.getHolidays();
	    final int finalYear = configWizardPanel.getSelectedYear();
	    final int finalMonth = configWizardPanel.getSelectedMonth();
	    
	    // Show loading indicator
	    configWizardPanel.showLoading();
	    configWizardPanel.setStatusMessage("Generating Monthly Report...", false);
	    
	    // Run report generation in background
	    SwingWorker<Boolean, Void> worker = new SwingWorker<Boolean, Void>() {
	        @Override
	        protected Boolean doInBackground() throws Exception {
	            try {
	                System.out.println("Generating Monthly Report...");
	                System.out.println("Report file: " + finalReportFilePath);
	                System.out.println("Merged file: " + finalMergedFilePath);
	                System.out.println("Year: " + finalYear + ", Month: " + finalMonth);
	                System.out.println("Holidays: " + finalHolidays);
	                
	                // Check if merged file exists
	                File mergedFile = new File(finalMergedFilePath);
	                if (!mergedFile.exists()) {
	                    System.err.println("Merged file not found: " + finalMergedFilePath);
	                    return false;
	                }
	                
	                // Generate monthly summary report
	                generateMonthlyReport(finalMergedFilePath, finalReportFilePath, finalHolidays, finalYear, finalMonth);
	                
	                return true;
	            } catch (Exception ex) {
	                ex.printStackTrace();
	                return false;
	            }
	        }
	        
	        @Override
	        protected void done() {
	            configWizardPanel.hideLoading();
	            
	            try {
	                boolean success = get();
	                if (success) {
	                    File reportFile = new File(finalReportFilePath);
	                    if (reportFile.exists()) {
	                        configWizardPanel.setStatusMessage(
	                            "✓ Monthly report created: " + reportFile.getAbsolutePath(), 
	                            false
	                        );
	                        
	                        // Generate and load data directly into ReportViewerPanel
	                        generateAndLoadMonthlyReport();
	                        
	                        // Ask user if they want to switch to Report panel
	                        int option = JOptionPane.showConfirmDialog(AttendanceAppGUI.this,
	                            "Monthly report generated successfully!\n\n" +
	                            "Do you want to switch to the Report panel to view the summary?",
	                            "Report Ready",
	                            JOptionPane.YES_NO_OPTION);
	                        
	                        if (option == JOptionPane.YES_OPTION) {
	                            showPanel("REPORT");
	                        }
	                        
	                        // Open the folder where the report is located
	                        try {
	                            Desktop.getDesktop().open(reportFile.getParentFile());
	                        } catch (Exception ex) {
	                            // Ignore if can't open folder
	                        }
	                    } else {
	                        configWizardPanel.setStatusMessage(
	                            "✓ Report generated but file not found at: " + finalReportFilePath, 
	                            false
	                        );
	                    }
	                } else {
	                    configWizardPanel.setStatusMessage(
	                        "✗ Failed to generate report. Check console for errors.", 
	                        true
	                    );
	                }
	            } catch (Exception ex) {
	                ex.printStackTrace();
	                configWizardPanel.setStatusMessage(
	                    "✗ Error: " + ex.getMessage(), 
	                    true
	                );
	            }
	        }
	    };
	    
	    worker.execute();
	}
	
	// New method to generate monthly summary report
	private void generateMonthlyReport(String mergedFilePath, String reportFilePath, 
	                                  List<Integer> holidays, int year, int month) throws Exception {
	    
	    AttendanceReportGenerator reportGenerator = new AttendanceReportGenerator();
	    
	    // Generate the report using CSV format with monthly summary
	    reportGenerator.exportToCsv(
	        reportFilePath.replace(".xlsx", ".csv"),  // Output CSV file
	        mergeResult.allEmployees,                 // Employee data
	        mergeResult.monthDays,                    // Number of days in month
	        holidays,                                 // Holiday list
	        year,                                     // Selected year
	        month                                     // Selected month
	    );
	    
	    // Also create Excel version for better formatting
	    try {
	        // Read the CSV and create Excel file
	        java.io.BufferedReader reader = new java.io.BufferedReader(new java.io.FileReader(reportFilePath.replace(".xlsx", ".csv")));
	        org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
	        org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("Monthly Summary");
	        
	        String line;
	        int rowNum = 0;
	        while ((line = reader.readLine()) != null) {
	            String[] columns = line.split(",");
	            org.apache.poi.ss.usermodel.Row row = sheet.createRow(rowNum++);
	            
	            for (int i = 0; i < columns.length; i++) {
	                org.apache.poi.ss.usermodel.Cell cell = row.createCell(i);
	                
	                // Try to parse as number
	                try {
	                    double value = Double.parseDouble(columns[i]);
	                    cell.setCellValue(value);
	                    
	                    // Format numeric cells
	                    if (i >= 1) { // All columns except Employee Code
	                        org.apache.poi.ss.usermodel.CellStyle numberStyle = workbook.createCellStyle();
	                        numberStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
	                        cell.setCellStyle(numberStyle);
	                    }
	                } catch (NumberFormatException e) {
	                    cell.setCellValue(columns[i]);
	                    
	                    // Style header row
	                    if (rowNum == 1) {
	                        org.apache.poi.ss.usermodel.CellStyle headerStyle = workbook.createCellStyle();
	                        org.apache.poi.ss.usermodel.Font headerFont = workbook.createFont();
	                        headerFont.setBold(true);
	                        headerStyle.setFont(headerFont);
	                        headerStyle.setFillForegroundColor(org.apache.poi.ss.usermodel.IndexedColors.GREY_25_PERCENT.getIndex());
	                        headerStyle.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
	                        cell.setCellStyle(headerStyle);
	                    }
	                }
	            }
	        }
	        
	        reader.close();
	        
	        // Auto-size columns
	        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
	            sheet.autoSizeColumn(i);
	        }
	        
	        // Write Excel file
	        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(reportFilePath)) {
	            workbook.write(fos);
	        }
	        
	        workbook.close();
	        
	    } catch (Exception e) {
	        System.err.println("Failed to create Excel version: " + e.getMessage());
	        // If Excel fails, at least CSV is created
	    }
	}
	
	private void exportMergedFile(ActionEvent e) {
		if (mergeResult == null) {
			configWizardPanel.setStatusMessage("Please process data first!", true);
			return;
		}

		// The merged file was already created by AttendanceMerger.merge()
		String mergedFilePath = configWizardPanel.getMergedFilePath();
		if (mergedFilePath == null || mergedFilePath.isEmpty()) {
			mergedFilePath = "Merged_Attendance.xlsx";
		}

		File mergedFile = new File(mergedFilePath);
		if (mergedFile.exists()) {
			// Open the folder where the file is located
			try {
				Desktop.getDesktop().open(mergedFile.getParentFile());
				configWizardPanel.setStatusMessage("Merged file location opened: " + mergedFile.getAbsolutePath(),
						false);
			} catch (Exception ex) {
				configWizardPanel.setStatusMessage("Merged file created at: " + mergedFile.getAbsolutePath(), false);
			}
		} else {
			configWizardPanel.setStatusMessage("Merged file not found at: " + mergedFilePath, true);
		}
	}

	private void queryData(ActionEvent e) {
		System.out.println("Querying data...");
	}

	private void exportFilteredData(ActionEvent e) {
	    if (dataViewerPanel.hasData()) {
	        // Create a file chooser
	        JFileChooser fileChooser = new JFileChooser();
	        fileChooser.setDialogTitle("Export Filtered Data");
	        fileChooser.setSelectedFile(new File("Filtered_Attendance_Data.xlsx"));
	        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
	            "Excel Files (*.xlsx)", "xlsx"));
	        
	        int result = fileChooser.showSaveDialog(this);
	        if (result == JFileChooser.APPROVE_OPTION) {
	            File selectedFile = fileChooser.getSelectedFile();
	            // Ensure .xlsx extension
	            if (!selectedFile.getName().toLowerCase().endsWith(".xlsx")) {
	                selectedFile = new File(selectedFile.getAbsolutePath() + ".xlsx");
	            }
	            
	            // Check if file exists
	            if (selectedFile.exists()) {
	                int overwrite = JOptionPane.showConfirmDialog(this,
	                    "File already exists. Overwrite?\n" + selectedFile.getName(),
	                    "Confirm Overwrite",
	                    JOptionPane.YES_NO_OPTION,
	                    JOptionPane.WARNING_MESSAGE);
	                
	                if (overwrite != JOptionPane.YES_OPTION) {
	                    return;
	                }
	            }
	            
	            // Export the data
	            exportTableToExcel(dataViewerPanel.getDataModel(), selectedFile);
	        }
	    } else {
	        JOptionPane.showMessageDialog(this, 
	            "No data to export. Please query data first.", 
	            "No Data", JOptionPane.WARNING_MESSAGE);
	    }
	}

	private void exportTableToExcel(DefaultTableModel model, File outputFile) {
	    try {
	        org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
	        org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("Filtered Data");
	        
	        // Create header row
	        org.apache.poi.ss.usermodel.Row headerRow = sheet.createRow(0);
	        for (int i = 0; i < model.getColumnCount(); i++) {
	            org.apache.poi.ss.usermodel.Cell cell = headerRow.createCell(i);
	            cell.setCellValue(model.getColumnName(i));
	            
	            // Style for header
	            org.apache.poi.ss.usermodel.CellStyle headerStyle = workbook.createCellStyle();
	            org.apache.poi.ss.usermodel.Font headerFont = workbook.createFont();
	            headerFont.setBold(true);
	            headerStyle.setFont(headerFont);
	            cell.setCellStyle(headerStyle);
	        }
	        
	        // Create data rows
	        for (int row = 0; row < model.getRowCount(); row++) {
	            org.apache.poi.ss.usermodel.Row dataRow = sheet.createRow(row + 1);
	            for (int col = 0; col < model.getColumnCount(); col++) {
	                org.apache.poi.ss.usermodel.Cell cell = dataRow.createCell(col);
	                Object value = model.getValueAt(row, col);
	                if (value != null) {
	                    cell.setCellValue(value.toString());
	                }
	            }
	        }
	        
	        // Auto-size columns
	        for (int i = 0; i < model.getColumnCount(); i++) {
	            sheet.autoSizeColumn(i);
	        }
	        
	        // Write to file
	        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(outputFile)) {
	            workbook.write(fos);
	        }
	        
	        workbook.close();
	        
	        JOptionPane.showMessageDialog(this,
	            "Data exported successfully to:\n" + outputFile.getAbsolutePath(),
	            "Export Complete",
	            JOptionPane.INFORMATION_MESSAGE);
	            
	    } catch (Exception ex) {
	        ex.printStackTrace();
	        JOptionPane.showMessageDialog(this,
	            "Error exporting data: " + ex.getMessage(),
	            "Export Error",
	            JOptionPane.ERROR_MESSAGE);
	    }
	}
	
	// Method to generate and load monthly report directly
	private void generateAndLoadMonthlyReport() {
	    if (mergeResult == null || reportViewerPanel == null) {
	        return;
	    }
	    
	    // Get configuration from config panel
	    final List<Integer> finalHolidays = configWizardPanel.getHolidays();
	    final int finalYear = configWizardPanel.getSelectedYear();
	    final int finalMonth = configWizardPanel.getSelectedMonth();
	    
	    // Show loading indicator
	    JDialog loadingDialog = new JDialog(this, "Generating Monthly Report", true);
	    loadingDialog.setLayout(new BorderLayout());
	    loadingDialog.setSize(300, 100);
	    loadingDialog.setLocationRelativeTo(this);
	    
	    JLabel loadingLabel = new JLabel("Generating monthly summary report...", SwingConstants.CENTER);
	    loadingLabel.setFont(new Font("Segoe UI", Font.PLAIN, 14));
	    JProgressBar progressBar = new JProgressBar();
	    progressBar.setIndeterminate(true);
	    
	    loadingDialog.add(loadingLabel, BorderLayout.CENTER);
	    loadingDialog.add(progressBar, BorderLayout.SOUTH);
	    
	    // Show loading in background thread
	    SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
	        @Override
	        protected Void doInBackground() throws Exception {
	            try {
	                List<String[]> reportData = new ArrayList<>();
	                List<String> headers = new ArrayList<>();
	                
	                // Add headers for monthly summary (one row per employee)
	                headers.add("Employee Code");
	                headers.add("Total Working Days");
	                headers.add("Total Full Days");
	                headers.add("Half Days Due To Duration");
	                headers.add("Half Days Due To PunchMiss");
	                headers.add("Half Days Due To Lates");
	                headers.add("Total Half Days");
	                headers.add("Total Lates");
	                headers.add("Total Absent");
	                headers.add("Total Punch Missed");
	                headers.add("Total OT Days");
	                headers.add("Total Half OT Days");
	                headers.add("Total OT Hours");
	                
	                // Calculate metrics for each employee
	                AttendanceReportGenerator generator = new AttendanceReportGenerator();
	                
	                for (AttendanceMerger.EmployeeData emp : mergeResult.allEmployees) {
	                    AttendanceReportGenerator.Metrics metrics = 
	                        generator.computeMetrics(emp, mergeResult.monthDays, finalHolidays, finalYear, finalMonth);
	                    
	                    String[] row = new String[headers.size()];
	                    int i = 0;
	                    row[i++] = emp.empId;
	                    row[i++] = String.valueOf(metrics.totalWorkingDays);
	                    row[i++] = String.valueOf(metrics.totalFullDays);
	                    row[i++] = String.valueOf(metrics.halfDaysDueToDuration);
	                    row[i++] = String.valueOf(metrics.halfDaysDueToPunchMiss);
	                    row[i++] = String.valueOf(metrics.halfDaysDueToLate);
	                    row[i++] = String.valueOf(metrics.halfDays);
	                    row[i++] = String.valueOf(metrics.totalLates);
	                    row[i++] = String.valueOf(metrics.totalAbsent);
	                    row[i++] = String.valueOf(metrics.totalPunchMissed);
	                    row[i++] = String.valueOf(metrics.totalOTDays);
	                    row[i++] = String.valueOf(metrics.totalHalfOTDays);
	                    
	                    // Calculate total OT hours (convert minutes to hours)
	                    double totalOTMinutes = metrics.workingDayOTMinutes + metrics.weekendFullOTMinutes + metrics.weekendHalfOTMinutes;
	                    double totalOTHours = Math.round((totalOTMinutes / 60.0) * 100.0) / 100.0;
	                    row[i++] = String.format("%.2f", totalOTHours);
	                    
	                    reportData.add(row);
	                }
	                
	                // Load into report panel on EDT
	                SwingUtilities.invokeLater(() -> {
	                    String[][] dataArray = reportData.toArray(new String[0][]);
	                    String[] headersArray = headers.toArray(new String[0]);
	                    reportViewerPanel.loadReportData(dataArray, headersArray);
	                    
	                    // Show success message
	                    JOptionPane.showMessageDialog(AttendanceAppGUI.this,
	                        String.format("Monthly Report Generated!\n\n" +
	                            "✓ %d employees processed\n" +
	                            "✓ One row per employee\n" +
	                            "✓ All monthly metrics included",
	                            mergeResult.allEmployees.size()),
	                        "Report Ready",
	                        JOptionPane.INFORMATION_MESSAGE);
	                });
	                
	            } catch (Exception e) {
	                e.printStackTrace();
	                SwingUtilities.invokeLater(() -> {
	                    JOptionPane.showMessageDialog(AttendanceAppGUI.this,
	                        "Error generating monthly report: " + e.getMessage(),
	                        "Report Error",
	                        JOptionPane.ERROR_MESSAGE);
	                });
	            }
	            return null;
	        }
	        
	        @Override
	        protected void done() {
	            loadingDialog.dispose();
	        }
	    };
	    
	    worker.execute();
	    loadingDialog.setVisible(true);
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> {
			try {
				new AttendanceAppGUI();
			} catch (Exception e) {
				e.printStackTrace();
				JOptionPane.showMessageDialog(null, "Error starting application: " + e.getMessage(), "Startup Error",
						JOptionPane.ERROR_MESSAGE);
			}
		});
	}
}