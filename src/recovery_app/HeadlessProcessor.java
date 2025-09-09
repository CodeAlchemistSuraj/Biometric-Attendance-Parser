package recovery_app;

import java.io.File;

public class HeadlessProcessor {
    public static void main(String[] args) {
        if (args == null || args.length == 0) {
            System.err.println("Usage: java recovery_app.HeadlessProcessor <excel-file-path>");
            System.exit(2);
        }
        String path = args[0];
        System.out.println("Headless processing: " + path);
        try {
            CleaningExcel_One processor = new CleaningExcel_One();
            processor.processWorkbook(new File(path));
            System.out.println("Processing completed: " + path);
        } catch (Exception e) {
            System.err.println("Processing failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}
