package org.bioparse.cleaning;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import com.formdev.flatlaf.FlatLightLaf; // Add this dependency for modern L&F

public class AttendanceAppGUI extends JFrame {

    private CardLayout cardLayout;
    private JPanel mainPanel, sidebarPanel;
    private AttendanceMerger.MergeResult mergeResult;
    private AttendanceReportGenerator reportGenerator = new AttendanceReportGenerator();
    private AttendanceQueryViewer queryViewer;

    // Fields from original (abbreviated for brevity)
    private JTextField inputFileField, mergedFileField /* etc. */;

    public AttendanceAppGUI() {
        try {
            UIManager.setLookAndFeel(new FlatLightLaf()); // Modern flat design
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
        menuBar.add(fileMenu);
        JMenu helpMenu = new JMenu("Help");
        helpMenu.add(new JMenuItem("About"));
        menuBar.add(helpMenu);
        setJMenuBar(menuBar);

        // Sidebar Navigation
        sidebarPanel = new JPanel(new GridLayout(4, 1, 0, 10));
        sidebarPanel.setPreferredSize(new Dimension(200, 0));
        sidebarPanel.setBackground(new Color(240, 240, 240)); // Light gray
        addButtonToSidebar("Home", e -> showPanel("HOME"));
        addButtonToSidebar("Configuration", e -> showPanel("CONFIG"));
        addButtonToSidebar("Data Viewer", e -> showPanel("DATA"));
        addButtonToSidebar("Report Viewer", e -> showPanel("REPORT"));
        add(sidebarPanel, BorderLayout.WEST);

        // Main Content with CardLayout
        cardLayout = new CardLayout();
        mainPanel = new JPanel(cardLayout);
        mainPanel.add(createHomePanel(), "HOME");
        mainPanel.add(createConfigWizardPanel(), "CONFIG"); // Wizard instead of single panel
        mainPanel.add(createDataViewerPanel(), "DATA");
        mainPanel.add(createReportViewerPanel(), "REPORT");
        add(mainPanel, BorderLayout.CENTER);

        showPanel("HOME"); // Start at dashboard
    }

    private void addButtonToSidebar(String text, java.awt.event.ActionListener listener) {
        JButton button = new JButton(text);
        button.addActionListener(listener);
        button.setFont(new Font("Arial", Font.BOLD, 14));
        button.setBackground(new Color(100, 181, 246)); // Blue accent
        button.setForeground(Color.WHITE);
        button.setToolTipText("Go to " + text);
        sidebarPanel.add(button);
    }

    private void showPanel(String panelName) {
        cardLayout.show(mainPanel, panelName);
    }

    // New Dashboard Home Panel
    private JPanel createHomePanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.add(new JLabel("Welcome to Attendance Parser Dashboard"), new GridBagConstraints(0, 0, 1, 1, 0, 0, GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(20, 0, 0, 0), 0, 0));
        // Add summary cards, e.g., recent files, quick stats (implement as needed)
        return panel;
    }

    // Config as Wizard (using CardLayout inside)
    private JPanel createConfigWizardPanel() {
        JPanel wizardPanel = new JPanel(new CardLayout());
        // Step 1: File Selection
        JPanel step1 = new JPanel();
        // Add fields like inputFileField, buttons
        JButton nextBtn = new JButton("Next");
        nextBtn.addActionListener(e -> ((CardLayout) wizardPanel.getLayout()).next(wizardPanel));
        step1.add(nextBtn);
        wizardPanel.add(step1, "STEP1");

        // Step 2: Holidays/Year/Month
        JPanel step2 = new JPanel();
        // Add holidaysField, etc.
        JButton backBtn = new JButton("Back");
        backBtn.addActionListener(e -> ((CardLayout) wizardPanel.getLayout()).previous(wizardPanel));
        JButton processBtn = new JButton("Process");
        processBtn.addActionListener(this::processData);
        step2.add(backBtn);
        step2.add(processBtn);
        wizardPanel.add(step2, "STEP2");

        return wizardPanel;
    }

    // Data Viewer and Report Viewer panels (from original, abbreviated)
    private JPanel createDataViewerPanel() {
        // Implement as in original
        return new JPanel();
    }

    private JPanel createReportViewerPanel() {
        // Implement as in original
        return new JPanel();
    }

    // Other methods (processData, queryData, etc.) remain similar

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new AttendanceAppGUI().setVisible(true));
    }
}