package org.bioparse;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bioparse.cleaning.CombinedDataCleanupAndEnhancement;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Scanner;

import static org.bioparse.cleaning.CombineSheetsToMasterClean.copyCellValue;

public class Main {
    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\MOHD GULZAR KHAN\\Desktop\\Suraj Working\\Java Git\\Biomteric Sheets\\outputv1.xlsx";
        String outputFilePath = "C:\\Users\\MOHD GULZAR KHAN\\Desktop\\Suraj Working\\Java Git\\output.xlsx";

        CombinedDataCleanupAndEnhancement processor = new CombinedDataCleanupAndEnhancement();



        try (FileInputStream fis = new FileInputStream(inputFilePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            // Step 1: Remove all images from all worksheets
            processor.removeImages(workbook);

            // Step 2: Remove blank rows from "WorkDuration" worksheet
            processor.removeBlankRowsFromWorkDuration(workbook);

            // Step 3: Process the main sheet (first sheet)
            processor.processMainSheet(workbook.getSheetAt(0));

            // Step 4: Add attendance metrics to the main sheet
            processor.addAttendanceMetrics(workbook.getSheetAt(0));

            // Save the modified workbook
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

            System.out.println("Processing complete!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
