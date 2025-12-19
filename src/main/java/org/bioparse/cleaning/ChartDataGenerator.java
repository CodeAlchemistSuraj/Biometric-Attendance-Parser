package org.bioparse.cleaning;

import org.jfree.chart.*;
import org.jfree.chart.plot.*;
import org.jfree.chart.renderer.category.*;
import org.jfree.chart.renderer.xy.*;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.PieDataset;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;
import org.jfree.chart.labels.*;
import org.jfree.chart.ui.TextAnchor;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.time.LocalTime;
import java.time.format.DateTimeParseException;
import java.time.Duration;
import java.util.*;

public class ChartDataGenerator {

    public static final Color[] CHART_COLORS = {
        new Color(70, 130, 180),    // Steel Blue
        new Color(60, 179, 113),    // Medium Sea Green
        new Color(255, 165, 0),     // Orange
        new Color(220, 20, 60),     // Crimson
        new Color(186, 85, 211),    // Medium Orchid
        new Color(46, 139, 87),     // Sea Green
        new Color(255, 99, 71),     // Tomato
        new Color(106, 90, 205)     // Slate Blue
    };

    // Compact base chart panel
    abstract static class BaseChartPanel extends JPanel {
        private static final long serialVersionUID = 1L;
        protected String title;
        
        public BaseChartPanel(String title) {
            this.title = title;
            setLayout(new BorderLayout(0, 5));
            setBackground(Color.WHITE);
            setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        }
        
        protected JPanel createExportPanel(JFreeChart chart) {
            JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 5, 0));
            panel.setOpaque(false);
            panel.setBorder(BorderFactory.createEmptyBorder(0, 0, 5, 0));
            
            JButton exportBtn = new JButton("📥");
            exportBtn.setToolTipText("Export as PNG");
            exportBtn.setPreferredSize(new Dimension(30, 25));
            exportBtn.setFont(new Font("Segoe UI", Font.PLAIN, 12));
            exportBtn.setFocusPainted(false);
            exportBtn.setBorder(BorderFactory.createEmptyBorder(2, 5, 2, 5));
            exportBtn.addActionListener(e -> exportChart(chart, "png"));
            
            panel.add(exportBtn);
            return panel;
        }
    }

    // Compact Bar Chart (200x200)
    abstract static class CompactBarChart extends BaseChartPanel {
        private static final long serialVersionUID = 1L;
        
        public CompactBarChart(String title, String xLabel, String yLabel) {
            super(title);
            createChart(xLabel, yLabel);
        }
        
        private void createChart(String xLabel, String yLabel) {
            DefaultCategoryDataset dataset = createDataset();
            JFreeChart chart = ChartFactory.createBarChart(
                title, xLabel, yLabel, dataset,
                PlotOrientation.VERTICAL, true, false, false
            );
            
            customizeChart(chart, dataset);
            
            ChartPanel chartPanel = new ChartPanel(chart);
            chartPanel.setPreferredSize(new Dimension(200, 200));
            chartPanel.setMinimumSize(new Dimension(180, 180));
            chartPanel.setMaximumDrawWidth(1000);
            chartPanel.setMaximumDrawHeight(1000);
            
            add(createExportPanel(chart), BorderLayout.NORTH);
            add(chartPanel, BorderLayout.CENTER);
        }
        
        private DefaultCategoryDataset createDataset() {
            DefaultCategoryDataset dataset = new DefaultCategoryDataset();
            Map<String, java.util.List<Double>> data = getData();
            java.util.List<String> categories = getCategories();
            
            for (Map.Entry<String, java.util.List<Double>> entry : data.entrySet()) {
                String series = entry.getKey();
                java.util.List<Double> values = entry.getValue();
                for (int i = 0; i < Math.min(values.size(), categories.size()); i++) {
                    dataset.addValue(values.get(i), series, categories.get(i));
                }
            }
            return dataset;
        }
        
        private void customizeChart(JFreeChart chart, CategoryDataset dataset) {
            CategoryPlot plot = chart.getCategoryPlot();
            BarRenderer renderer = (BarRenderer) plot.getRenderer();
            
            // Use gradient colors
            for (int i = 0; i < dataset.getRowCount() && i < CHART_COLORS.length; i++) {
                Color base = CHART_COLORS[i % CHART_COLORS.length];
                renderer.setSeriesPaint(i, base);
            }
            
            // Remove item labels for compactness
            renderer.setDefaultItemLabelsVisible(false);
            
            // Compact styling
            plot.setBackgroundPaint(new Color(250, 250, 250));
            plot.setRangeGridlinePaint(new Color(220, 220, 220));
            plot.setDomainGridlinePaint(new Color(220, 220, 220));
            plot.setOutlineVisible(false);
            
            // Compact fonts
            chart.getTitle().setFont(new Font("Segoe UI", Font.BOLD, 14));
            plot.getDomainAxis().setLabelFont(new Font("Segoe UI", Font.PLAIN, 11));
            plot.getRangeAxis().setLabelFont(new Font("Segoe UI", Font.PLAIN, 11));
            plot.getDomainAxis().setTickLabelFont(new Font("Segoe UI", Font.PLAIN, 9));
            plot.getRangeAxis().setTickLabelFont(new Font("Segoe UI", Font.PLAIN, 9));
            
            // Reduce padding
            plot.getDomainAxis().setCategoryMargin(0.15);
            renderer.setItemMargin(0.15);
        }
        
        protected abstract Map<String, java.util.List<Double>> getData();
        protected abstract java.util.List<String> getCategories();
    }

    // Compact Pie Chart (200x200)
    abstract static class CompactPieChart extends BaseChartPanel {
        private static final long serialVersionUID = 1L;
        
        public CompactPieChart(String title) {
            super(title);
            createChart();
        }
        
        private void createChart() {
            DefaultPieDataset<String> dataset = createDataset();
            JFreeChart chart = ChartFactory.createPieChart(
                title, dataset, true, false, false
            );
            
            customizeChart(chart, dataset);
            
            ChartPanel chartPanel = new ChartPanel(chart);
            chartPanel.setPreferredSize(new Dimension(200, 200));
            chartPanel.setMinimumSize(new Dimension(180, 180));
            
            add(createExportPanel(chart), BorderLayout.NORTH);
            add(chartPanel, BorderLayout.CENTER);
        }
        
        private DefaultPieDataset<String> createDataset() {
            DefaultPieDataset<String> dataset = new DefaultPieDataset<>();
            Map<String, Double> data = getData();
            // Sort by value descending
            data.entrySet().stream()
                .sorted(Map.Entry.<String, Double>comparingByValue().reversed())
                .forEachOrdered(entry -> dataset.setValue(entry.getKey(), entry.getValue()));
            return dataset;
        }
        
        private void customizeChart(JFreeChart chart, PieDataset<String> dataset) {
            PiePlot<String> plot = (PiePlot<String>) chart.getPlot();
            
            // Apply colors
            for (int i = 0; i < dataset.getItemCount() && i < CHART_COLORS.length; i++) {
                plot.setSectionPaint(dataset.getKey(i), CHART_COLORS[i % CHART_COLORS.length]);
            }
            
            // Compact label settings
            plot.setLabelGenerator(new StandardPieSectionLabelGenerator("{1}"));
            plot.setLabelFont(new Font("Segoe UI", Font.PLAIN, 10));
            plot.setLabelBackgroundPaint(new Color(255, 255, 255, 200));
            plot.setLabelShadowPaint(null);
            plot.setLabelOutlinePaint(null);
            
            plot.setBackgroundPaint(new Color(250, 250, 250));
            plot.setOutlineVisible(false);
            plot.setShadowPaint(null);
            
            chart.getTitle().setFont(new Font("Segoe UI", Font.BOLD, 14));
            
            // Show legend instead of labels if too many items
            if (dataset.getItemCount() > 5) {
                plot.setLabelGenerator(null);
            }
        }
        
        protected abstract Map<String, Double> getData();
    }

    // Compact Line Chart (200x200)
    abstract static class CompactLineChart extends BaseChartPanel {
        private static final long serialVersionUID = 1L;
        
        public CompactLineChart(String title, String xLabel, String yLabel) {
            super(title);
            createChart(xLabel, yLabel);
        }
        
        private void createChart(String xLabel, String yLabel) {
            XYSeriesCollection dataset = createDataset();
            JFreeChart chart = ChartFactory.createXYLineChart(
                title, xLabel, yLabel, dataset,
                PlotOrientation.VERTICAL, true, false, false
            );
            
            customizeChart(chart);
            
            ChartPanel chartPanel = new ChartPanel(chart);
            chartPanel.setPreferredSize(new Dimension(200, 200));
            chartPanel.setMinimumSize(new Dimension(180, 180));
            
            add(createExportPanel(chart), BorderLayout.NORTH);
            add(chartPanel, BorderLayout.CENTER);
        }
        
        private XYSeriesCollection createDataset() {
            XYSeriesCollection dataset = new XYSeriesCollection();
            Map<String, java.util.List<Double>> data = getData();
            
            for (Map.Entry<String, java.util.List<Double>> entry : data.entrySet()) {
                XYSeries series = new XYSeries(entry.getKey());
                java.util.List<Double> values = entry.getValue();
                for (int i = 0; i < values.size(); i++) {
                    series.add(i, values.get(i));
                }
                dataset.addSeries(series);
            }
            return dataset;
        }
        
        private void customizeChart(JFreeChart chart) {
            XYPlot plot = chart.getXYPlot();
            XYLineAndShapeRenderer renderer = new XYLineAndShapeRenderer();
            
            for (int i = 0; i < plot.getSeriesCount() && i < CHART_COLORS.length; i++) {
                Color color = CHART_COLORS[i % CHART_COLORS.length];
                renderer.setSeriesPaint(i, color);
                renderer.setSeriesStroke(i, new BasicStroke(1.5f));
                renderer.setSeriesShapesVisible(i, true);
                renderer.setSeriesShape(i, new java.awt.geom.Ellipse2D.Double(-2, -2, 4, 4));
            }
            
            plot.setRenderer(renderer);
            plot.setBackgroundPaint(new Color(250, 250, 250));
            plot.setRangeGridlinePaint(new Color(220, 220, 220));
            plot.setDomainGridlinePaint(new Color(220, 220, 220));
            plot.setOutlineVisible(false);
            
            chart.getTitle().setFont(new Font("Segoe UI", Font.BOLD, 14));
            plot.getDomainAxis().setLabelFont(new Font("Segoe UI", Font.PLAIN, 11));
            plot.getRangeAxis().setLabelFont(new Font("Segoe UI", Font.PLAIN, 11));
            plot.getDomainAxis().setTickLabelFont(new Font("Segoe UI", Font.PLAIN, 9));
            plot.getRangeAxis().setTickLabelFont(new Font("Segoe UI", Font.PLAIN, 9));
        }
        
        protected abstract Map<String, java.util.List<Double>> getData();
    }

    // Export functionality
    private static void exportChart(JFreeChart chart, String format) {
        JFileChooser fc = new JFileChooser();
        fc.setSelectedFile(new File(chart.getTitle().getText() + "." + format));
        if (fc.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            try {
                File file = fc.getSelectedFile();
                if ("png".equalsIgnoreCase(format)) {
                    ChartUtils.saveChartAsPNG(file, chart, 600, 400);
                    JOptionPane.showMessageDialog(null, "Chart exported: " + file.getName(),
                        "Export Successful", JOptionPane.INFORMATION_MESSAGE);
                }
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Export failed: " + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    // ==================== COMPACT CHART METHODS ====================

    public static JPanel createAttendanceTrendChart(java.util.List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new CompactBarChart("Daily Trend", "Day", "Count") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, java.util.List<Double>> getData() {
                Map<String, java.util.List<Double>> data = new LinkedHashMap<>();
                data.put("P", new java.util.ArrayList<>());
                data.put("A", new java.util.ArrayList<>());
                
                // Sample every 3 days for compactness
                for (int day = 1; day <= monthDays; day += 3) {
                    int present = 0, absent = 0;
                    for (AttendanceMerger.EmployeeData emp : employees) {
                        String status = AttendanceUtils.safeGet(emp.dailyData.get("Status"), day - 1);
                        if ("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status) || "WOP".equalsIgnoreCase(status)) {
                            present++;
                        } else if ("A".equalsIgnoreCase(status)) absent++;
                    }
                    data.get("P").add((double) present);
                    data.get("A").add((double) absent);
                }
                return data;
            }
            
            @Override 
            protected java.util.List<String> getCategories() {
                java.util.List<String> cats = new java.util.ArrayList<>();
                // Show every 3rd day label
                for (int i = 1; i <= monthDays; i += 3) {
                    cats.add("D" + i);
                }
                return cats;
            }
        };
    }

    public static JPanel createAttendanceDistributionChart(java.util.List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new CompactPieChart("Attendance") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, Double> getData() {
                int present = 0, absent = 0, half = 0, leave = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    java.util.List<String> statusList = emp.dailyData.get("Status");
                    if (statusList != null) {
                        for (String s : statusList) {
                            if (s == null || s.trim().isEmpty()) continue;
                            switch (s.toUpperCase()) {
                                case "P": case "WO": case "WOP": present++; break;
                                case "A": absent++; break;
                                case "HD": case "H": half++; break;
                                case "L": leave++; break;
                                default: present++;
                            }
                        }
                    }
                }
                Map<String, Double> data = new LinkedHashMap<>();
                data.put("Present", (double) present);
                data.put("Absent", (double) absent);
                data.put("Half Day", (double) half);
                data.put("Leave", (double) leave);
                return data;
            }
        };
    }

    public static JPanel createDepartmentWiseAttendanceChart(java.util.List<AttendanceMerger.EmployeeData> employees) {
        return new CompactBarChart("By Department", "Dept", "Count") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, java.util.List<Double>> getData() {
                Map<String, Integer> deptCounts = new LinkedHashMap<>();
                String[] depts = {"IT", "HR", "Finance", "Ops", "Admin"};
                
                for (String dept : depts) deptCounts.put(dept, 0);
                
                for (AttendanceMerger.EmployeeData emp : employees) {
                    String dept = depts[Math.abs(emp.empId.hashCode()) % depts.length];
                    deptCounts.put(dept, deptCounts.get(dept) + 1);
                }
                
                Map<String, java.util.List<Double>> data = new LinkedHashMap<>();
                data.put("Emp", new java.util.ArrayList<>());
                deptCounts.values().forEach(v -> data.get("Emp").add(v.doubleValue()));
                return data;
            }
            
            @Override 
            protected java.util.List<String> getCategories() {
                return Arrays.asList("IT", "HR", "Finance", "Ops", "Admin");
            }
        };
    }

    public static JPanel createTopPerformersChart(java.util.List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new CompactBarChart("Performance", "Range", "Count") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, java.util.List<Double>> getData() {
                int excellent = 0, good = 0, average = 0, poor = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    int present = 0;
                    java.util.List<String> statusList = emp.dailyData.get("Status");
                    if (statusList != null) {
                        for (String s : statusList) {
                            if ("P".equalsIgnoreCase(s) || "WO".equalsIgnoreCase(s) || "WOP".equalsIgnoreCase(s)) present++;
                        }
                    }
                    double rate = (present * 100.0) / monthDays;
                    if (rate >= 90) excellent++;
                    else if (rate >= 75) good++;
                    else if (rate >= 60) average++;
                    else poor++;
                }
                Map<String, java.util.List<Double>> data = new LinkedHashMap<>();
                data.put("Emp", Arrays.asList((double)excellent, (double)good, (double)average, (double)poor));
                return data;
            }
            
            @Override 
            protected java.util.List<String> getCategories() {
                return Arrays.asList("90%+", "75-89%", "60-74%", "<60%");
            }
        };
    }

    public static JPanel createLateArrivalsAnalysis(java.util.List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new CompactBarChart("Late Arrivals", "Delay", "Count") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, java.util.List<Double>> getData() {
                int[] ranges = new int[4];
                for (AttendanceMerger.EmployeeData emp : employees) {
                    java.util.List<String> inTimes = emp.dailyData.get("InTime");
                    java.util.List<String> statusList = emp.dailyData.get("Status");
                    if (inTimes == null || statusList == null) continue;
                    for (int i = 0; i < Math.min(inTimes.size(), statusList.size()); i++) {
                        String status = statusList.get(i);
                        String inTime = inTimes.get(i);
                        if (!("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status) || "WOP".equalsIgnoreCase(status))) continue;
                        if (!isLateArrival(inTime)) continue;
                        try {
                            LocalTime arrival = LocalTime.parse(inTime);
                            LocalTime threshold = LocalTime.parse("09:45");
                            long minutesLate = Duration.between(threshold, arrival).toMinutes();
                            if (minutesLate <= 15) ranges[0]++;
                            else if (minutesLate <= 30) ranges[1]++;
                            else if (minutesLate <= 60) ranges[2]++;
                            else ranges[3]++;
                        } catch (DateTimeParseException ignored) {}
                    }
                }
                Map<String, java.util.List<Double>> data = new LinkedHashMap<>();
                data.put("Late", Arrays.asList((double)ranges[0], (double)ranges[1], (double)ranges[2], (double)ranges[3]));
                return data;
            }
            
            @Override 
            protected java.util.List<String> getCategories() {
                return Arrays.asList("0-15", "16-30", "31-60", "60+");
            }
        };
    }

    public static JPanel createOvertimeAnalysisChart(java.util.List<AttendanceMerger.EmployeeData> employees) {
        return new CompactPieChart("Overtime") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, Double> getData() {
                int none = 0, low = 0, medium = 0, high = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    double totalOT = 0;
                    java.util.List<String> durations = emp.dailyData.get("Duration");
                    if (durations != null) {
                        for (String d : durations) {
                            if (d == null || d.isEmpty() || "-".equals(d)) continue;
                            try {
                                String[] parts = d.split(":");
                                int h = Integer.parseInt(parts[0]);
                                int m = parts.length > 1 ? Integer.parseInt(parts[1]) : 0;
                                double mins = h * 60 + m;
                                if (mins > 510) totalOT += (mins - 510) / 60.0;
                            } catch (Exception ignored) {}
                        }
                    }
                    if (totalOT == 0) none++;
                    else if (totalOT <= 10) low++;
                    else if (totalOT <= 30) medium++;
                    else high++;
                }
                Map<String, Double> data = new LinkedHashMap<>();
                data.put("None", (double)none);
                data.put("1-10h", (double)low);
                data.put("11-30h", (double)medium);
                data.put("30+h", (double)high);
                return data;
            }
        };
    }

    public static JPanel createDailyAttendanceLineChart(java.util.List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new CompactLineChart("Daily Trend", "Day", "Count") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, java.util.List<Double>> getData() {
                java.util.List<Double> present = new java.util.ArrayList<>();
                java.util.List<Double> absent = new java.util.ArrayList<>();
                
                // Sample data points (reduce for compactness)
                for (int day = 1; day <= monthDays; day += 2) {
                    int p = 0, a = 0;
                    for (AttendanceMerger.EmployeeData emp : employees) {
                        String status = AttendanceUtils.safeGet(emp.dailyData.get("Status"), day - 1);
                        if ("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status) || "WOP".equalsIgnoreCase(status)) p++;
                        else if ("A".equalsIgnoreCase(status)) a++;
                    }
                    present.add((double)p); 
                    absent.add((double)a);
                }
                
                Map<String, java.util.List<Double>> data = new LinkedHashMap<>();
                data.put("Present", present);
                data.put("Absent", absent);
                return data;
            }
        };
    }

    public static JPanel createWeeklyAttendanceChart(java.util.List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new CompactBarChart("Weekly Pattern", "Week", "Days") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, java.util.List<Double>> getData() {
                Map<String, java.util.List<Double>> data = new LinkedHashMap<>();
                data.put("P", new java.util.ArrayList<>());
                data.put("A", new java.util.ArrayList<>());
                
                int weeks = (monthDays + 6) / 7;
                for (int w = 0; w < weeks; w++) {
                    int present = 0, absent = 0;
                    int start = w * 7 + 1, end = Math.min(start + 6, monthDays);
                    for (int d = start; d <= end; d++) {
                        for (AttendanceMerger.EmployeeData emp : employees) {
                            String status = AttendanceUtils.safeGet(emp.dailyData.get("Status"), d - 1);
                            if ("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status) || "WOP".equalsIgnoreCase(status)) {
                                present++;
                            } else if ("A".equalsIgnoreCase(status)) absent++;
                        }
                    }
                    data.get("P").add((double)present);
                    data.get("A").add((double)absent);
                }
                return data;
            }
            
            @Override 
            protected java.util.List<String> getCategories() {
                java.util.List<String> cats = new java.util.ArrayList<>();
                int weeks = (monthDays + 6) / 7;
                for (int i = 1; i <= weeks; i++) cats.add("W" + i);
                return cats;
            }
        };
    }

    public static JPanel createShiftComplianceChart(java.util.List<AttendanceMerger.EmployeeData> employees) {
        return new CompactPieChart("Punctuality") {
            private static final long serialVersionUID = 1L;
            
            @Override 
            protected Map<String, Double> getData() {
                int onTime = 0, late = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    java.util.List<String> inTimes = emp.dailyData.get("InTime");
                    java.util.List<String> statusList = emp.dailyData.get("Status");
                    if (inTimes == null || statusList == null) continue;
                    for (int i = 0; i < Math.min(inTimes.size(), statusList.size()); i++) {
                        String status = statusList.get(i);
                        String inTime = inTimes.get(i);
                        if (!("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status))) continue;
                        if (isLateArrival(inTime)) late++;
                        else onTime++;
                    }
                }
                Map<String, Double> data = new LinkedHashMap<>();
                data.put("On Time", (double)onTime);
                data.put("Late", (double)late);
                return data;
            }
        };
    }

    // Utility method for late arrival detection
    private static boolean isLateArrival(String inTime) {
        if (inTime == null || inTime.isEmpty() || "-".equals(inTime)) return false;
        try {
            LocalTime arrival = LocalTime.parse(inTime);
            LocalTime threshold = LocalTime.parse("09:45");
            return arrival.isAfter(threshold);
        } catch (DateTimeParseException e) {
            return false;
        }
    }
}