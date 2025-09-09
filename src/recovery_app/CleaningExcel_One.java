package recovery_app;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class CleaningExcel_One extends Application {

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Excel Data Cleanup and Enhancement Tool");

        VBox root = new VBox(10);
        root.setPadding(new Insets(20));

        Button processButton = new Button("Select and Process Excel File");
        processButton.setOnAction(e -> processExcelFile(primaryStage));

        Label statusLabel = new Label("Ready to process...");

        root.getChildren().addAll(processButton, statusLabel);

        Scene scene = new Scene(root, 400, 200);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void processExcelFile(Stage stage) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select Excel File");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx", "*.xls"));
        File selectedFile = fileChooser.showOpenDialog(stage);

        if (selectedFile != null) {
            try {
                processWorkbook(selectedFile);
                showAlert(Alert.AlertType.INFORMATION, "Success", "Processing completed successfully!");
            } catch (Exception ex) {
                showAlert(Alert.AlertType.ERROR, "Error", "Failed to process file: " + ex.getMessage());
                ex.printStackTrace();
            }
        }
    }

    private void showAlert(Alert.AlertType type, String title, String content) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setContentText(content);
        alert.showAndWait();
    }

    private void processWorkbook(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            // Step 1: Remove all images from all worksheets
            for (Sheet sheet : workbook) {
                if (sheet instanceof XSSFSheet) {
                    XSSFSheet xssfSheet = (XSSFSheet) sheet;
                    // Remove pictures by clearing the drawing patriarch
                    XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
                    if (drawing != null) {
                        // POI doesn't support direct removal of all pictures easily.
                        // Workaround: Create a new drawing patriarch to clear existing pictures.
                        xssfSheet.getCTWorksheet().unsetDrawing();
                        // Re-create an empty drawing patriarch if needed
                        xssfSheet.createDrawingPatriarch();
                    }
                }
                // Note: For HSSFSheet (.xls), image removal is not well-supported by POI.
                // If .xls support is needed, consider copying content to a new sheet without images.
            }

            // Step 2: Remove blank rows from "WorkDuration" worksheet
            Sheet workDurationSheet = null;
            for (Sheet sheet : workbook) {
                if ("WorkDuration".equals(sheet.getSheetName())) {
                    workDurationSheet = sheet;
                    break;
                }
            }
            if (workDurationSheet != null) {
                removeBlankRows(workDurationSheet);
            }

            // Step 3: Process the first sheet as active sheet (assuming active is index 0)
            Sheet activeSheet = workbook.getSheetAt(0);
            processActiveSheet(activeSheet);

            // Save the modified workbook
            String outputPath = file.getAbsolutePath().replace(".xlsx", "_processed.xlsx");
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
            }

            System.out.println("Processed file saved to: " + outputPath);
        }
    }

    private void processActiveSheet(Sheet sheet) {
        // Unprotect the sheet if protected (assuming no password)
        if (sheet.getProtect()) {
            // Note: POI requires a password to unprotect if one is set.
            // If a password exists, you'll need to provide it or handle it differently.
            // For simplicity, assume no password or warn user if protected.
            // Since POI doesn't have a direct unprotect method without password, we skip if password-protected.
            // Alternatively, you can prompt the user for a password or log a warning.
            System.out.println("Warning: Sheet is protected. Unprotection skipped as password handling is not implemented.");
        }

        // Set row height to 12 (in points)
        for (Row row : sheet) {
            if (row != null) {
                row.setHeightInPoints(12);
            }
        }

        int lastRowNum = sheet.getLastRowNum();
        List<Integer> rowsToDelete = new ArrayList<>();

        // Remove rows with "Department:" in column A
        for (int currentRow = lastRowNum; currentRow >= 0; currentRow--) {
            Row row = sheet.getRow(currentRow);
            if (row != null) {
                Cell cellA = row.getCell(0);
                if (cellA != null && cellA.getCellType() == CellType.STRING) {
                    if (cellA.getStringCellValue().toLowerCase().startsWith("department:")) {
                        rowsToDelete.add(currentRow);
                    }
                }
            }
        }
        // Delete rows in reverse order to avoid index shifting issues
        for (int rowIdx : rowsToDelete) {
            Row row = sheet.getRow(rowIdx);
            if (row != null) {
                sheet.removeRow(row);
                // Shift rows up to fill the gap
                if (rowIdx < lastRowNum) {
                    sheet.shiftRows(rowIdx + 1, lastRowNum, -1);
                }
                lastRowNum--;
            }
        }
        lastRowNum = updateLastRow(sheet);

        // Insert empty row before each "Employee:" row
        rowsToDelete.clear();
        for (int currentRow = 0; currentRow <= lastRowNum; currentRow++) {
            Row row = sheet.getRow(currentRow);
            if (row != null) {
                Cell cellA = row.getCell(0);
                if (cellA != null && cellA.getCellType() == CellType.STRING) {
                    if (cellA.getStringCellValue().toLowerCase().startsWith("employee:")) {
                        // Insert row above
                        sheet.shiftRows(currentRow, lastRowNum, 1, true, false);
                        // Create empty row at currentRow
                        Row emptyRow = sheet.createRow(currentRow);
                        lastRowNum++;
                    }
                }
            }
        }
        lastRowNum = updateLastRow(sheet);

        // Process employee blocks to remove "Late By", "Early By", "OT"
        int currentRow = 1; // Start from row 1 (0-indexed row 1 is 2nd row)
        while (currentRow <= lastRowNum) {
            Row row = sheet.getRow(currentRow);
            if (row != null && hasTextStartingWith(row.getCell(0), "employee:")) {
                int blockStart = currentRow + 1;
                int blockEnd = Math.min(blockStart + 7, lastRowNum + 1);
                Set<Integer> deleteRows = new HashSet<>();
                if (blockStart + 4 < blockEnd) deleteRows.add(blockStart + 4); // Late By
                if (blockStart + 5 < blockEnd) deleteRows.add(blockStart + 5); // Early By
                if (blockStart + 6 < blockEnd) deleteRows.add(blockStart + 6); // OT
                for (int delRow : deleteRows) {
                    if (delRow <= lastRowNum) {
                        Row delRowObj = sheet.getRow(delRow);
                        if (delRowObj != null) {
                            sheet.removeRow(delRowObj);
                            if (delRow < lastRowNum) {
                                sheet.shiftRows(delRow + 1, lastRowNum, -1);
                            }
                            lastRowNum--;
                        }
                    }
                }
                currentRow = blockStart + 8 - deleteRows.size();
            } else {
                currentRow++;
            }
        }
        lastRowNum = updateLastRow(sheet);

        // Step 4: Add FullDay, HalfDay, OTDays, OTHours, HalfOTDay rows
        int lastColNum = getLastColumn(sheet);
        currentRow = 1; // Status likely starts around here

        while (currentRow <= lastRowNum) {
            Row row = sheet.getRow(currentRow);
            if (row != null && hasText(row.getCell(0), "Status")) {
                int blockStartRow = currentRow;
                int statusRow = currentRow;
                int inTimeRow = currentRow + 1;
                int outTimeRow = currentRow + 2;
                int durationRow = currentRow + 3;
                int shiftRow = currentRow + 4;
                int blockEndRow = currentRow + 4;

                // Find blockLastCol
                int blockLastCol = findBlockLastCol(sheet, statusRow, blockEndRow, lastColNum);

                // Insert 5 rows after blockEndRow
                sheet.shiftRows(blockEndRow + 1, lastRowNum + 1, 5, true, false);
                lastRowNum += 5;

                // Create and label new rows (0-indexed)
                int fullDayRow = blockEndRow + 1;
                int halfDayRow = blockEndRow + 2;
                int otDaysRow = blockEndRow + 3;
                int otHoursRow = blockEndRow + 4;
                int halfOtDayRow = blockEndRow + 5;

                Row fullDayR = sheet.createRow(fullDayRow);
                fullDayR.createCell(0).setCellValue("FullDay");
                Row halfDayR = sheet.createRow(halfDayRow);
                halfDayR.createCell(0).setCellValue("HalfDay");
                Row otDaysR = sheet.createRow(otDaysRow);
                otDaysR.createCell(0).setCellValue("OTDays");
                Row otHoursR = sheet.createRow(otHoursRow);
                otHoursR.createCell(0).setCellValue("OTHours");
                Row halfOtDayR = sheet.createRow(halfOtDayRow);
                halfOtDayR.createCell(0).setCellValue("HalfOTDay");

                // Apply borders to the block (dotted thin black)
                CellStyle borderStyle = sheet.getWorkbook().createCellStyle();
                borderStyle.setBorderTop(BorderStyle.DOTTED);
                borderStyle.setBorderBottom(BorderStyle.DOTTED);
                borderStyle.setBorderLeft(BorderStyle.DOTTED);
                borderStyle.setBorderRight(BorderStyle.DOTTED);
                borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
                borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

                // Apply to range
                for (int r = statusRow; r <= halfOtDayRow; r++) {
                    Row blockRow = sheet.getRow(r);
                    if (blockRow == null) blockRow = sheet.createRow(r);
                    for (int c = 0; c <= blockLastCol; c++) {
                        Cell cell = blockRow.getCell(c);
                        if (cell == null) cell = blockRow.createCell(c);
                        cell.setCellStyle(borderStyle);
                    }
                }

                // Apply formulas from column 1 (B, 0-index 1) to blockLastCol
                for (int col = 1; col <= blockLastCol; col++) {
                    String daysCell = columnLetter(col + 1) + "1"; // Row 1, col+1 since 0-index

                    // FullDay Formula
                    String fullDayFormula = "=IF(AND(" + cellAddress(statusRow, col) + "=\"P\", " +
                            cellAddress(inTimeRow, col) + "<>\"\", " +
                            cellAddress(outTimeRow, col) + "<>\"\", " +
                            cellAddress(outTimeRow, col) + "-" + cellAddress(inTimeRow, col) + ">=TIME(8,30,0), " +
                            "AND(ISERROR(SEARCH(\"St\"," + daysCell + ")), ISERROR(SEARCH(\"S\"," + daysCell + ")))), 1, 0)";
                    fullDayR.createCell(col).setCellFormula(fullDayFormula);

                    // HalfDay Formula
                    String halfDayFormula = "=IF(AND(" + cellAddress(fullDayRow, col) + "=0, " +
                            cellAddress(inTimeRow, col) + "<>\"\", " +
                            cellAddress(outTimeRow, col) + "<>\"\", " +
                            cellAddress(outTimeRow, col) + "-" + cellAddress(inTimeRow, col) + ">=TIME(4,0,0), " +
                            "OR(" + cellAddress(statusRow, col) + "=\"P\", " +
                            cellAddress(statusRow, col) + "=\"½P\")), 0.5, 0)";
                    halfDayR.createCell(col).setCellFormula(halfDayFormula);

                    // OTDays Formula
                    String otDaysFormula = "=IF(OR(" + cellAddress(statusRow, col) + "=\"WO\", " +
                            cellAddress(statusRow, col) + "=\"H\", " +
                            cellAddress(statusRow, col) + "=\"HP\", " +
                            cellAddress(statusRow, col) + "=\"WOP\"), " +
                            "IF(AND(" + cellAddress(inTimeRow, col) + "<>\"\", " +
                            cellAddress(outTimeRow, col) + "<>\"\"), 1, " +
                            "IF(OR(" + cellAddress(inTimeRow, col) + "<>\"\", " +
                            cellAddress(outTimeRow, col) + "<>\"\"), 0.5, 0)), 0)";
                    otDaysR.createCell(col).setCellFormula(otDaysFormula);

                    // OTHours Formula
                    String otHoursFormula = "=IF(AND(" + cellAddress(statusRow, col) + "=\"P\", " +
                            "NOT(ISBLANK(" + cellAddress(inTimeRow, col) + ")), " +
                            "NOT(ISBLANK(" + cellAddress(outTimeRow, col) + "))), " +
                            "MAX(0, (" + cellAddress(outTimeRow, col) + " - " + cellAddress(inTimeRow, col) + ") - (8.5/24)), 0)";
                    otHoursR.createCell(col).setCellFormula(otHoursFormula);

                    // HalfOTDay Formula
                    String halfOtDayFormula = "=IF(OR(" + cellAddress(statusRow, col) + "=\"WO\", " +
                            cellAddress(statusRow, col) + "=\"H\", " +
                            cellAddress(statusRow, col) + "=\"HP\", " +
                            cellAddress(statusRow, col) + "=\"WOP\"), " +
                            "IF(XOR(ISBLANK(" + cellAddress(inTimeRow, col) + "), " +
                            "ISBLANK(" + cellAddress(outTimeRow, col) + ")), 1, 0), 0)";
                    halfOtDayR.createCell(col).setCellFormula(halfOtDayFormula);
                }

                currentRow = blockEndRow + 6;
            } else if (isEmptyCell(sheet.getRow(currentRow), 0) || hasTextStartingWith(sheet.getRow(currentRow) != null ? sheet.getRow(currentRow).getCell(0) : null, "employee:")) {
                currentRow++;
            } else {
                currentRow++;
            }
        }

        // Auto-fit columns
        for (int i = 0; i <= lastColNum; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void removeBlankRows(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = lastRowNum; i >= 0; i--) {
            Row row = sheet.getRow(i);
            if (row == null || isRowBlank(row)) {
                if (row != null) {
                    sheet.removeRow(row);
                }
                if (i < lastRowNum) {
                    sheet.shiftRows(i + 1, lastRowNum, -1);
                }
                lastRowNum--;
            }
        }
    }

    private boolean isRowBlank(Row row) {
        if (row == null) return true;
        for (int j = 0; j < row.getLastCellNum(); j++) {
            Cell cell = row.getCell(j);
            if (cell != null && cell.getCellType() != CellType.BLANK && !isEmpty(cell)) {
                return false;
            }
        }
        return true;
    }

    private boolean isEmpty(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK ||
               (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty());
    }

    private int updateLastRow(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        // Ensure lastRowNum is accurate by checking for non-empty rows
        while (lastRowNum >= 0 && (sheet.getRow(lastRowNum) == null || isRowBlank(sheet.getRow(lastRowNum)))) {
            lastRowNum--;
        }
        return lastRowNum;
    }

    private int getLastColumn(Sheet sheet) {
        int maxCol = 0;
        for (Row row : sheet) {
            if (row != null && row.getLastCellNum() > maxCol) {
                maxCol = row.getLastCellNum();
            }
        }
        return maxCol - 1; // 0-index
    }

    private int findBlockLastCol(Sheet sheet, int startRow, int endRow, int maxCol) {
        int lastCol = 0;
        for (int col = 1; col <= maxCol; col++) { // Start from B
            boolean hasData = false;
            for (int r = startRow; r <= endRow; r++) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    Cell cell = row.getCell(col);
                    if (cell != null && !isEmpty(cell)) {
                        hasData = true;
                        break;
                    }
                }
            }
            if (hasData) {
                lastCol = col;
            }
        }
        return lastCol;
    }

    private boolean hasText(Cell cell, String text) {
        return cell != null && cell.getCellType() == CellType.STRING &&
               cell.getStringCellValue().trim().equalsIgnoreCase(text);
    }

    private boolean hasTextStartingWith(Cell cell, String prefix) {
        return cell != null && cell.getCellType() == CellType.STRING &&
               cell.getStringCellValue().toLowerCase().startsWith(prefix.toLowerCase());
    }

    private boolean isEmptyCell(Row row, int col) {
        Cell cell = row != null ? row.getCell(col) : null;
        return cell == null || cell.getCellType() == CellType.BLANK ||
               (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty());
    }

    private String cellAddress(int row, int col) {
        return columnLetter(col + 1) + (row + 1);
    }

    private String columnLetter(int colNum) {
        StringBuilder sb = new StringBuilder();
        while (colNum > 0) {
            int remainder = (colNum - 1) % 26;
            sb.insert(0, (char) ('A' + remainder));
            colNum = (colNum - 1) / 26;
        }
        return sb.toString();
    }

    public static void main(String[] args) {
        launch(args);
    }
}