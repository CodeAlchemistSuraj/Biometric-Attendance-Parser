# RecoveryWorkApp

This repository contains a small Java application to clean and enhance Excel workbooks.

There are two ways to run the tool:

1. Headless (recommended if JavaFX GUI fails on your system)

   This mode runs the processing logic directly from the command line and does not require JavaFX.

   Example:

   ```powershell
   cd 'C:\Users\MOHD GULZAR KHAN\eclipse-workspace\RecoveryWorkApp'
   javac -cp "lib/*;." -d bin src\recovery_app\*.java
   java -cp "bin;lib/*" recovery_app.HeadlessProcessor "C:\path\to\your\file.xlsx"
   ```

   The processed file will be saved next to the input file with `_processed.xlsx` appended to the name.

2. GUI (JavaFX)

   If you prefer the JavaFX UI, run with JavaFX jars on the module path. JavaFX requires a compatible graphics pipeline.

   Recommended command (module mode):

   ```powershell
   cd 'C:\Users\MOHD GULZAR KHAN\eclipse-workspace\RecoveryWorkApp'
   java -Dprism.order=sw -Dprism.forceGPU=false -Dprism.verbose=true --module-path "lib" --add-modules javafx.controls,javafx.fxml -cp "bin;lib/*" recovery_app.CleaningExcel_One
   ```

   If you see warnings about native access (JDK strictness), try adding `--enable-native-access=javafx.graphics` (some JVMs require this). Example ordering may matter:

   ```powershell
   java --module-path lib --add-modules javafx.controls,javafx.fxml --enable-native-access=javafx.graphics -Dprism.order=sw -Dprism.forceGPU=false -Dprism.verbose=true -cp "bin;lib/*" recovery_app.CleaningExcel_One
   ```

   Alternate classpath-based run (may work on some systems):

   ```powershell
   java -Dprism.order=sw -Dprism.forceGPU=false -Dprism.verbose=true -Djavafx.platform=monocle -Dmonocle.platform=Software -cp "bin;lib/*" recovery_app.CleaningExcel_One
   ```

Troubleshooting tips
- Make sure the JavaFX jars in `lib/` match your JDK version (major version should match, e.g., JavaFX 24 for Java 24).
- If you get `Error initializing QuantumRenderer: no suitable pipeline found`, try the software rendering mode (`-Dprism.order=sw`) and use the headless processor if GUI still fails.
- For headless or CI usage, use `HeadlessProcessor` which processes a workbook without starting JavaFX.

If you want, I can also add a small example Excel file or create a packaged runnable script for Windows.
# RecoveryWorkApp