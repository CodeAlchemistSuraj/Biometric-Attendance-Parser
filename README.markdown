Biometric Attendance Parser
Overview
Biometric Attendance Parser is a Java-based desktop application designed to process biometric attendance data from Excel files. It merges multiple attendance sheets into a single master file, generates employee reports with key metrics (e.g., lates, leaves, overtime), and provides a user-friendly GUI for querying and viewing attendance data. The app skips weekends and holidays, computes duration-based metrics (full days, half days), and supports visual displays via tables.
The application is built using:

Apache POI for Excel file manipulation.
FlatLaf for a modern UI look and feel.
Swing for the graphical user interface.

This tool is ideal for HR departments or businesses managing employee attendance from biometric systems, automating manual tasks and providing actionable insights.
Features

Data Merging: Merge multiple biometric Excel sheets into a master file, detecting month/year and handling holidays/weekends.
Report Generation: Compute employee metrics such as total lates, leaves, full days, half days, overtime days/hours, punch misses, and half OT days. Outputs to Excel or CSV.
Query and View: Query attendance by employee code, date, or date range, displaying results in tables or text summaries.
GUI with Navigation: Sidebar navigation for Home, Configuration, Data Viewer, and Report Viewer. Configuration uses a wizard-style flow for ease of use.
Error Handling: Built-in validation and error messages for file selection, processing, and queries.
Customizable: Supports holiday input, year/month selection, and multi-file input.

Requirements

Java: JDK 17 (LTS version, recommended for stability).
Maven: 3.9.11 or later for building the project.
Dependencies (handled by Maven):

Apache POI 5.3.0 for Excel processing.
FlatLaf 3.2.5 for UI styling.
Commons Logging 1.3.0 for logging.


OS: Windows 10/11, Linux, or macOS (tested on Windows 10).
Input Data: Excel files (.xlsx) with biometric attendance data (e.g., monthly status reports with Employee, InTime, OutTime, Duration, Status).

Installation

Clone the Repository:
textgit clone https://your-repo-url/Biometric-Attendance-Parser.git
cd Biometric-Attendance-Parser

Set Up JAVA_HOME:

Download JDK 17 from Adoptium or Oracle.
Set environment variable:

Windows: set JAVA_HOME=C:\Program Files\Java\jdk-17 (add to Path: %JAVA_HOME%\bin).
Linux/macOS: export JAVA_HOME=/usr/lib/jvm/java-17-openjdk (add to .bashrc or .zshrc).


Verify:
textjava -version



Build with Maven:
textmvn clean install

This compiles the code and generates a JAR file in target/Biometric-Attendance-Parser-1.0-SNAPSHOT.jar.


Eclipse Setup (Optional for Development):

Import the project as a Maven project.
Configure JRE: Window > Preferences > Java > Installed JREs > Add JDK 17.
Update build path: Right-click project > Build Path > Configure Build Path > Remove unbound JRE > Add JDK 17.
Maven > Update Project.



How to Run

Run the JAR File:
textjava -jar target/Biometric-Attendance-Parser-1.0-SNAPSHOT.jar

The GUI will launch with a sidebar navigation.


Using the GUI:

Home: Dashboard with welcome message and recent summaries (extendable for stats).
Configuration (Wizard):

Step 1: Choose input Excel files (multi-select) and merged output file.
Step 2: Enter holidays (comma-separated, e.g., "15,26"), select year and month.
Click "Process" to merge data and compute metrics.


Data Viewer: Enter employee code, day, or date range to query attendance. Results show in tables.
Report Viewer: Click "Generate Report" to view or save metrics (lates, leaves, OT, etc.).
File Menu: Open input, save report.
Help Menu: About dialog.


Command-Line Alternative (Optional):

Use AttendanceProcessor.java (if available) for CLI mode:
textjava -cp target/Biometric-Attendance-Parser-1.0-SNAPSHOT.jar org.bioparse.cleaning.AttendanceProcessor

Hardcode input/output paths in the code for batch processing.



Example Data
Before Processing (Input Data)
Input: Multiple Excel files (.xlsx) with raw biometric attendance data. Example structure for a single sheet:

Header: "August 15 2025 Monthly Status Report"
Employee block:

Employee: EMP001 : John Doe
InTime: (cells for days 1-31)
OutTime: (cells for days 1-31)
Duration: (e.g., "08:30")
Status: (e.g., "P", "L" for leave)



Example snippet from input Excel:















































12...31Employee: EMP001 : John DoeInTime09:0009:15...09:00OutTime17:3017:45...17:30Duration08:3008:30...08:30StatusPP...L

Holidays: 15, 26 (e.g., Independence Day, Republic Day).
Weekends skipped automatically.

After Processing (Output Data)

Merged File (Merged.xlsx): Consolidated master sheet with employee blocks, skipping holidays/weekends.
Example:




















Employee: EMP001 : John DoeInTimeOutTimeDurationStatus

Report File (Employee_Report.xlsx): Summary metrics.
Example:































Employee CodeTotal Working DaysTotal LatesTotal LeavesTotal Full DaysHalf DaysFinal HalfDaysTotal OT DaysTotal OT HoursTotal Punch MissedTotal Half OT DaysEMP001253220421210

Query Example:

By Employee EMP001: Table with daily InTime, OutTime, Duration, Status.
By Date 1: List of all employees' data for day 1.
By Date Range 1-5: Detailed table for the range.



Value Delivered to Business

Efficiency: Automates merging and report generation from raw biometric data, reducing manual effort from hours to minutes for HR teams.
Accuracy: Automatically skips weekends/holidays, calculates duration-based metrics (full/half days, OT), and detects punch misses, minimizing errors in payroll and attendance tracking.
Insights: Provides key metrics (lates, leaves, OT) for performance reviews, compliance (e.g., labor laws), and resource planning.
User-Friendly: GUI with wizard and viewers simplifies usage for non-technical users, with export to Excel/CSV for sharing.
Cost Savings: Reduces administrative costs by streamlining attendance analysis, enabling better workforce management.
Scalability: Handles multiple files and large datasets, suitable for small to medium businesses.

Data Analysis
The app performs the following analysis on attendance data:

Metric Calculation:

Total Working Days: Non-holiday/weekend days with duration or status.
Total Lates: InTime after 09:15.
Total Leaves: Status "L".
Total Full Days: Duration ≥ 510 minutes (8h 30m).
Half Days: Duration ≥ 240 minutes (4h) but < 510 minutes; final half days = half days / 2.
Total OT Days/Hours: Duration > 510 + 30 minutes; OT hours = (extra minutes) / 60.
Total Punch Missed: Duration > 0 but InTime or OutTime missing.
Total Half OT Days: OT < 240 minutes.


Data Processing:

Merges sheets, detects month/year, trims to month days (e.g., 31 for August).
Skips weekends (Sat/Sun) and holidays.
Parses durations (HH:MM to minutes) for thresholds.


Query Analysis: Filters by employee/date/range, displaying raw data (InTime, OutTime, Duration, Status).
Insights: Enables trend analysis (e.g., frequent lates), compliance checks, and performance metrics for business decisions.

If you encounter any issues or need customization, open an issue on the repository or contact the developer.

This README is comprehensive and self-contained. If you need additions (e.g., screenshots, license), let me know!
