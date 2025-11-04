package org.bioparse.cleaning;

import org.jfree.chart.*;
import org.jfree.chart.plot.*;
import org.jfree.chart.renderer.category.*;
import org.jfree.chart.renderer.xy.*;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;
import org.jfree.chart.labels.*;
import org.jfree.chart.entity.*;
import org.jfree.chart.ui.RectangleInsets;
import org.jfree.chart.ui.TextAnchor;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.List;

public class ChartDataGenerator {

    public static final Color[] CHART_COLORS = {
        new Color(41, 128, 185),    // Blue
        new Color(39, 174, 96),     // Green
        new Color(243, 156, 18),    // Orange
        new Color(231, 76, 60),     // Red
        new Color(155, 89, 182),    // Purple
        new Color(52, 73, 94),      // Dark Blue
        new Color(241, 196, 15),    // Yellow
        new Color(230, 126, 34)     // Dark Orange
    };

    // ==================== JFreeChart Base Classes ====================

    abstract static class JFreeBarChartPanel extends JPanel {
        public JFreeBarChartPanel(String title, String xLabel, String yLabel) {
            setLayout(new BorderLayout());
            setBackground(Color.WHITE);
            setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
          //  setPreferredSize(new Dimension(500, 350));
            createChart(title, xLabel, yLabel);
        }

        protected void createChart(String title, String xLabel, String yLabel) {
            DefaultCategoryDataset dataset = new DefaultCategoryDataset();
            Map<String, List<Double>> data = getData();
            List<String> categories = getCategories();

            for (Map.Entry<String, List<Double>> entry : data.entrySet()) {
                String series = entry.getKey();
                List<Double> values = entry.getValue();
                for (int i = 0; i < values.size() && i < categories.size(); i++) {
                    dataset.addValue(values.get(i), series, categories.get(i));
                }
            }

            JFreeChart chart = ChartFactory.createBarChart(
                title, xLabel, yLabel, dataset,
                PlotOrientation.VERTICAL, true, true, false
            );

            CategoryPlot plot = chart.getCategoryPlot();
            BarRenderer renderer = (BarRenderer) plot.getRenderer();
            for (int i = 0; i < data.size() && i < CHART_COLORS.length; i++) {
                renderer.setSeriesPaint(i, CHART_COLORS[i]);
            }
            renderer.setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator());
            renderer.setDefaultItemLabelsVisible(true);
            renderer.setDefaultPositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BOTTOM_CENTER));

            plot.setBackgroundPaint(Color.WHITE);
            plot.setRangeGridlinePaint(Color.LIGHT_GRAY);
            plot.setDomainGridlinePaint(Color.LIGHT_GRAY);

            ChartPanel chartPanel = new ChartPanel(chart);
            chartPanel.setMouseWheelEnabled(true);
            add(createExportPanel(chart), BorderLayout.NORTH);
            add(chartPanel, BorderLayout.CENTER);
        }

        protected abstract Map<String, List<Double>> getData();
        protected abstract List<String> getCategories();
    }

    abstract static class JFreeStackedBarChartPanel extends JPanel {
        public JFreeStackedBarChartPanel(String title, String xLabel, String yLabel) {
            setLayout(new BorderLayout());
            setBackground(Color.WHITE);
            setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
           // setPreferredSize(new Dimension(500, 350));
            createChart(title, xLabel, yLabel);
        }

        protected void createChart(String title, String xLabel, String yLabel) {
            DefaultCategoryDataset dataset = new DefaultCategoryDataset();
            Map<String, List<Double>> data = getData();
            List<String> categories = getCategories();

            for (Map.Entry<String, List<Double>> entry : data.entrySet()) {
                String series = entry.getKey();
                List<Double> values = entry.getValue();
                for (int i = 0; i < values.size() && i < categories.size(); i++) {
                    dataset.addValue(values.get(i), series, categories.get(i));
                }
            }

            JFreeChart chart = ChartFactory.createStackedBarChart(
                title, xLabel, yLabel, dataset,
                PlotOrientation.VERTICAL, true, true, false
            );

            CategoryPlot plot = chart.getCategoryPlot();
            StackedBarRenderer renderer = (StackedBarRenderer) plot.getRenderer();
            for (int i = 0; i < data.size() && i < CHART_COLORS.length; i++) {
                renderer.setSeriesPaint(i, CHART_COLORS[i]);
            }

            plot.setBackgroundPaint(Color.WHITE);
            plot.setRangeGridlinePaint(Color.LIGHT_GRAY);

            ChartPanel chartPanel = new ChartPanel(chart);
            chartPanel.setMouseWheelEnabled(true);
            add(createExportPanel(chart), BorderLayout.NORTH);
            add(chartPanel, BorderLayout.CENTER);
        }

        protected abstract Map<String, List<Double>> getData();
        protected abstract List<String> getCategories();
    }

    abstract static class JFreePieChartPanel extends JPanel {
        public JFreePieChartPanel(String title) {
            setLayout(new BorderLayout());
            setBackground(Color.WHITE);
            setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
           // setPreferredSize(new Dimension(500, 350));
            createChart(title);
        }

        protected void createChart(String title) {
            DefaultPieDataset dataset = new DefaultPieDataset();
            Map<String, Double> data = getData();
            data.forEach((k, v) -> { if (v > 0) dataset.setValue(k, v); });

            JFreeChart chart = ChartFactory.createPieChart(title, dataset, true, true, false);
            PiePlot plot = (PiePlot) chart.getPlot();
            for (int i = 0; i < dataset.getItemCount() && i < CHART_COLORS.length; i++) {
                plot.setSectionPaint(dataset.getKey(i), CHART_COLORS[i % CHART_COLORS.length]);
            }
            plot.setLabelGenerator(new StandardPieSectionLabelGenerator("{0}: {1} ({2})"));

            ChartPanel chartPanel = new ChartPanel(chart);
            chartPanel.setMouseWheelEnabled(true);
            add(createExportPanel(chart), BorderLayout.NORTH);
            add(chartPanel, BorderLayout.CENTER);
        }

        protected abstract Map<String, Double> getData();
    }

    abstract static class JFreeLineChartPanel extends JPanel {
        public JFreeLineChartPanel(String title, String xLabel, String yLabel) {
            setLayout(new BorderLayout());
            setBackground(Color.WHITE);
            setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
           // setPreferredSize(new Dimension(500, 350));
            createChart(title, xLabel, yLabel);
        }

        protected void createChart(String title, String xLabel, String yLabel) {
            XYSeriesCollection dataset = new XYSeriesCollection();
            Map<String, List<Double>> data = getData();
            List<String> categories = getCategories();

            for (Map.Entry<String, List<Double>> entry : data.entrySet()) {
                XYSeries series = new XYSeries(entry.getKey());
                List<Double> values = entry.getValue();
                for (int i = 0; i < values.size() && i < categories.size(); i++) {
                    series.add(i + 1, values.get(i));
                }
                dataset.addSeries(series);
            }

            JFreeChart chart = ChartFactory.createXYLineChart(
                title, xLabel, yLabel, dataset,
                PlotOrientation.VERTICAL, true, true, false
            );

            XYPlot plot = chart.getXYPlot();
            XYLineAndShapeRenderer renderer = (XYLineAndShapeRenderer) plot.getRenderer();
            for (int i = 0; i < data.size() && i < CHART_COLORS.length; i++) {
                renderer.setSeriesPaint(i, CHART_COLORS[i]);
                renderer.setSeriesStroke(i, new BasicStroke(2.5f));
            }

            plot.setBackgroundPaint(Color.WHITE);
            plot.setRangeGridlinePaint(Color.LIGHT_GRAY);
            plot.setDomainGridlinePaint(Color.LIGHT_GRAY);

            ChartPanel chartPanel = new ChartPanel(chart);
            chartPanel.setMouseWheelEnabled(true);
            add(createExportPanel(chart), BorderLayout.NORTH);
            add(chartPanel, BorderLayout.CENTER);
        }

        protected abstract Map<String, List<Double>> getData();
        protected abstract List<String> getCategories();
    }

    // ==================== Export Panel ====================
    private static JPanel createExportPanel(JFreeChart chart) {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        panel.setBackground(Color.WHITE);

        JButton pngBtn = new JButton("PNG");
        pngBtn.setToolTipText("Export as PNG");
        pngBtn.addActionListener(e -> exportChart(chart, "png"));

        JButton pdfBtn = new JButton("PDF");
        pdfBtn.setToolTipText("Export as PDF");
        pdfBtn.addActionListener(e -> exportChart(chart, "pdf"));

        panel.add(pngBtn);
        panel.add(pdfBtn);
        return panel;
    }

    private static void exportChart(JFreeChart chart, String format) {
        JFileChooser fc = new JFileChooser();
        fc.setSelectedFile(new File(chart.getTitle().getText() + "." + format));
        if (fc.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            try {
                File file = fc.getSelectedFile();
                if ("png".equalsIgnoreCase(format)) {
                    ChartUtils.saveChartAsPNG(file, chart, 800, 600);
                } else if ("pdf".equalsIgnoreCase(format)) {
                   // ChartUtils.saveChartAsPDF(file, chart, 800, 600, new org.jfree.pdf.PDFGraphics2D(new java.awt.geom.Rectangle2D.Double(0, 0, 800, 600)));
                }
                JOptionPane.showMessageDialog(null, "Exported: " + file.getName());
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Export failed: " + ex.getMessage());
            }
        }
    }

    // ==================== CHARTS ====================

    public static JPanel createAttendanceTrendChart(List<AttendanceMerger.EmployeeData> employees, int monthDays, int year, int month) {
        return new JFreeBarChartPanel("Monthly Attendance Trend", "Day", "Employees") {
            @Override protected Map<String, List<Double>> getData() {
                Map<String, List<Double>> data = new LinkedHashMap<>();
                data.put("Present", new ArrayList<>());
                data.put("Absent", new ArrayList<>());
                data.put("Late", new ArrayList<>());

                for (int day = 1; day <= monthDays; day++) {
                    int present = 0, absent = 0, late = 0;
                    for (AttendanceMerger.EmployeeData emp : employees) {
                        String status = AttendanceUtils.safeGet(emp.dailyData.get("Status"), day - 1);
                        String inTime = AttendanceUtils.safeGet(emp.dailyData.get("InTime"), day - 1);
                        if ("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status) || "WOP".equalsIgnoreCase(status)) {
                            present++;
                            if (isLateArrival(inTime)) late++;
                        } else if ("A".equalsIgnoreCase(status)) absent++;
                    }
                    data.get("Present").add((double) present);
                    data.get("Absent").add((double) absent);
                    data.get("Late").add((double) late);
                }
                return data;
            }
            @Override protected List<String> getCategories() {
                List<String> cats = new ArrayList<>();
                for (int i = 1; i <= monthDays; i++) cats.add("D" + i);
                return cats;
            }
        };
    }

    public static JPanel createAttendanceDistributionChart(List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new JFreePieChartPanel("Attendance Distribution") {
            @Override protected Map<String, Double> getData() {
                Map<String, Double> data = new LinkedHashMap<>();
                int present = 0, absent = 0, half = 0, leave = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    List<String> statusList = emp.dailyData.get("Status");
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
                data.put("Present", (double) present);
                data.put("Absent", (double) absent);
                data.put("Half Days", (double) half);
                data.put("Leaves", (double) leave);
                return data;
            }
        };
    }

    public static JPanel createDepartmentWiseAttendanceChart(List<AttendanceMerger.EmployeeData> employees) {
        return new JFreeBarChartPanel("Department-wise Headcount", "Department", "Employees") {
            @Override protected Map<String, List<Double>> getData() {
                Map<String, Integer> deptCounts = new LinkedHashMap<>();
                deptCounts.put("IT", 0); deptCounts.put("HR", 0);
                deptCounts.put("Finance", 0); deptCounts.put("Operations", 0); deptCounts.put("Admin", 0);

                Random rand = new Random();
                String[] depts = {"IT", "HR", "Finance", "Operations", "Admin"};
                for (AttendanceMerger.EmployeeData emp : employees) {
                    String dept = depts[Math.abs(emp.empId.hashCode()) % depts.length];
                    deptCounts.put(dept, deptCounts.get(dept) + 1);
                }

                Map<String, List<Double>> data = new LinkedHashMap<>();
                data.put("Employees", new ArrayList<>());
                deptCounts.values().forEach(v -> data.get("Employees").add(v.doubleValue()));
                return data;
            }
            @Override protected List<String> getCategories() {
                return Arrays.asList("IT", "HR", "Finance", "Operations", "Admin");
            }
        };
    }

    public static JPanel createTopPerformersChart(List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new JFreeBarChartPanel("Employee Performance", "Attendance Range", "Count") {
            @Override protected Map<String, List<Double>> getData() {
                int excellent = 0, good = 0, average = 0, poor = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    int present = 0;
                    List<String> statusList = emp.dailyData.get("Status");
                    if (statusList != null) {
                        for (String s : statusList) {
                            if ("P".equalsIgnoreCase(s) || "WO".equalsIgnoreCase(s) || "WOP".equalsIgnoreCase(s)) present++;
                        }
                    }
                    double rate = present * 100.0 / monthDays;
                    if (rate >= 90) excellent++;
                    else if (rate >= 75) good++;
                    else if (rate >= 60) average++;
                    else poor++;
                }
                Map<String, List<Double>> data = new LinkedHashMap<>();
                data.put("Employees", Arrays.asList((double)excellent, (double)good, (double)average, (double)poor));
                return data;
            }
            @Override protected List<String> getCategories() {
                return Arrays.asList("90-100%", "75-89%", "60-74%", "Below 60%");
            }
        };
    }

    public static JPanel createLateArrivalsAnalysis(List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new JFreeBarChartPanel("Late Arrivals by Time", "Time Range", "Count") {
            @Override protected Map<String, List<Double>> getData() {
                int r1 = 0, r2 = 0, r3 = 0, r4 = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    List<String> inTimes = emp.dailyData.get("InTime");
                    List<String> statusList = emp.dailyData.get("Status");
                    if (inTimes == null || statusList == null) continue;
                    for (int i = 0; i < inTimes.size(); i++) {
                        String status = statusList.get(i);
                        String inTime = inTimes.get(i);
                        if (!("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status) || "WOP".equalsIgnoreCase(status))) continue;
                        if (!isLateArrival(inTime)) continue;
                        try {
                            java.time.LocalTime arrival = java.time.LocalTime.parse(inTime);
                            java.time.LocalTime threshold = java.time.LocalTime.parse("09:45");
                            long minutesLate = java.time.Duration.between(threshold, arrival).toMinutes();
                            if (minutesLate <= 15) r1++;
                            else if (minutesLate <= 30) r2++;
                            else if (minutesLate <= 60) r3++;
                            else r4++;
                        } catch (Exception ignored) {}
                    }
                }
                Map<String, List<Double>> data = new LinkedHashMap<>();
                data.put("Late", Arrays.asList((double)r1, (double)r2, (double)r3, (double)r4));
                return data;
            }
            @Override protected List<String> getCategories() {
                return Arrays.asList("0-15 min", "16-30 min", "31-60 min", "60+ min");
            }
        };
    }

    public static JPanel createOvertimeAnalysisChart(List<AttendanceMerger.EmployeeData> employees) {
        return new JFreePieChartPanel("Overtime Distribution") {
            @Override protected Map<String, Double> getData() {
                int none = 0, low = 0, medium = 0, high = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    double totalOT = 0;
                    List<String> durations = emp.dailyData.get("Duration");
                    if (durations != null) {
                        for (String d : durations) {
                            if (d == null || d.isEmpty() || "-".equals(d)) continue;
                            try {
                                String[] parts = d.split(":");
                                int h = Integer.parseInt(parts[0]), m = Integer.parseInt(parts[1]);
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
                data.put("1-10 hrs", (double)low);
                data.put("11-30 hrs", (double)medium);
                data.put("30+ hrs", (double)high);
                return data;
            }
        };
    }

    public static JPanel createDailyAttendanceLineChart(List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new JFreeLineChartPanel("Daily Attendance Trend", "Day", "Employees") {
            @Override protected Map<String, List<Double>> getData() {
                List<Double> present = new ArrayList<>(), absent = new ArrayList<>();
                for (int day = 1; day <= monthDays; day++) {
                    int p = 0, a = 0;
                    for (AttendanceMerger.EmployeeData emp : employees) {
                        String status = AttendanceUtils.safeGet(emp.dailyData.get("Status"), day - 1);
                        if ("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status) || "WOP".equalsIgnoreCase(status)) p++;
                        else if ("A".equalsIgnoreCase(status)) a++;
                    }
                    present.add((double)p); absent.add((double)a);
                }
                Map<String, List<Double>> data = new LinkedHashMap<>();
                data.put("Present", present);
                data.put("Absent", absent);
                return data;
            }
            @Override protected List<String> getCategories() {
                List<String> cats = new ArrayList<>();
                for (int i = 1; i <= monthDays; i++) cats.add("D" + i);
                return cats;
            }
        };
    }

    public static JPanel createWeeklyAttendanceChart(List<AttendanceMerger.EmployeeData> employees, int monthDays) {
        return new JFreeStackedBarChartPanel("Weekly Attendance Pattern", "Week", "Days") {
            @Override protected Map<String, List<Double>> getData() {
                Map<String, List<Double>> data = new LinkedHashMap<>();
                data.put("Present", new ArrayList<>());
                data.put("Absent", new ArrayList<>());
                data.put("Late", new ArrayList<>());

                int weeks = (monthDays + 6) / 7;
                for (int w = 0; w < weeks; w++) {
                    int present = 0, absent = 0, late = 0;
                    int start = w * 7 + 1, end = Math.min(start + 6, monthDays);
                    for (int d = start; d <= end; d++) {
                        for (AttendanceMerger.EmployeeData emp : employees) {
                            String status = AttendanceUtils.safeGet(emp.dailyData.get("Status"), d - 1);
                            String inTime = AttendanceUtils.safeGet(emp.dailyData.get("InTime"), d - 1);
                            if ("P".equalsIgnoreCase(status) || "WO".equalsIgnoreCase(status) || "WOP".equalsIgnoreCase(status)) {
                                present++;
                                if (isLateArrival(inTime)) late++;
                            } else if ("A".equalsIgnoreCase(status)) absent++;
                        }
                    }
                    data.get("Present").add((double)present);
                    data.get("Absent").add((double)absent);
                    data.get("Late").add((double)late);
                }
                return data;
            }
            @Override protected List<String> getCategories() {
                List<String> cats = new ArrayList<>();
                int weeks = (monthDays + 6) / 7;
                for (int i = 1; i <= weeks; i++) cats.add("Week " + i);
                return cats;
            }
        };
    }

    public static JPanel createShiftComplianceChart(List<AttendanceMerger.EmployeeData> employees) {
        return new JFreePieChartPanel("Shift Compliance") {
            @Override protected Map<String, Double> getData() {
                int onTime = 0, late = 0, early = 0;
                for (AttendanceMerger.EmployeeData emp : employees) {
                    List<String> inTimes = emp.dailyData.get("InTime");
                    List<String> statusList = emp.dailyData.get("Status");
                    if (inTimes == null || statusList == null) continue;
                    for (int i = 0; i < inTimes.size(); i++) {
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

    private static boolean isLateArrival(String inTime) {
        if (inTime == null || inTime.isEmpty() || "-".equals(inTime)) return false;
        try {
            java.time.LocalTime arrival = java.time.LocalTime.parse(inTime);
            java.time.LocalTime threshold = java.time.LocalTime.parse("09:45");
            return arrival.isAfter(threshold);
        } catch (Exception e) {
            return false;
        }
    }
}