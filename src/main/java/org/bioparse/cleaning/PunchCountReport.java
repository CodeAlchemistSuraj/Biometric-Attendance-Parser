package org.bioparse.cleaning;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class PunchCountReport {

    public static void main(String[] args) throws Exception {
        String inputFilePath = "C:\\Users\\MOHD GULZAR KHAN\\Desktop\\Suraj Working\\Java Git\\Biomteric Sheets\\Attendance.xlsx";
        String outputFilePath = "C:\\Users\\MOHD GULZAR KHAN\\Desktop\\Suraj Working\\Java Git\\Biomteric Sheets\\DevicePunchReportv4.xlsx";

        generatePunchCountReport(inputFilePath, outputFilePath);
    }

    public static void generatePunchCountReport(String inputFile, String outputFile) throws Exception {
        FileInputStream fis = new FileInputStream(inputFile);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // employeeCode -> (deviceName -> count)
        Map<String, Map<String, Integer>> dataMap = new LinkedHashMap<>();
        String currentEmpCode = null;

        for (Row row : sheet) {
            if (row == null || row.getLastCellNum() == -1) continue;

            int lastCol = row.getLastCellNum();
            String[] cells = new String[lastCol];
            for (int i = 0; i < lastCol; i++) {
                Cell cell = row.getCell(i);
                cells[i] = (cell == null) ? "" : cell.toString().trim();
            }

            if (cells[0].equalsIgnoreCase("Employee")) {
                if (cells.length > 2 && cells[2].contains(":")) {
                    String[] parts = cells[2].split(":");
                    if (parts.length > 0) {
                        currentEmpCode = parts[0].trim();
                        dataMap.putIfAbsent(currentEmpCode, new LinkedHashMap<>());
                    }
                }
            } else if (cells[0].matches("\\d{2}-[A-Za-z]{3}-\\d{4}.*") && currentEmpCode != null) {
                String device = (cells.length > 2) ? cells[2].trim() : "";
                if (!device.isEmpty()) {
                    dataMap.putIfAbsent(currentEmpCode, new LinkedHashMap<>());
                    Map<String, Integer> devMap = dataMap.get(currentEmpCode);
                    devMap.put(device, devMap.getOrDefault(device, 0) + 1);
                }
            }
        }

        workbook.close();
        fis.close();

        writeOutputExcel(outputFile, dataMap);
        System.out.println("✅ Report generated: " + outputFile);
    }

    private static void writeOutputExcel(String filePath, Map<String, Map<String, Integer>> data) throws Exception {
        Workbook outWb = new XSSFWorkbook();
        Sheet outSheet = outWb.createSheet("Punch Report");

        // Collect all unique device names
        Set<String> devices = new TreeSet<>();
        for (Map<String, Integer> empData : data.values()) {
            devices.addAll(empData.keySet());
        }

        // Header
        Row header = outSheet.createRow(0);
        header.createCell(0).setCellValue("Employee Code");
        int col = 1;
        for (String device : devices) {
            header.createCell(col++).setCellValue(device);
        }

        // Data rows
        int rowNum = 1;
        for (String empCode : data.keySet()) {
            Row row = outSheet.createRow(rowNum++);
            row.createCell(0).setCellValue(empCode);
            Map<String, Integer> devMap = data.get(empCode);
            col = 1;
            for (String device : devices) {
                row.createCell(col++).setCellValue(devMap.getOrDefault(device, 0));
            }
        }

        // Auto-size columns
        for (int i = 0; i <= devices.size(); i++) {
            outSheet.autoSizeColumn(i);
        }

        FileOutputStream fos = new FileOutputStream(filePath);
        outWb.write(fos);
        fos.close();
        outWb.close();
    }
}