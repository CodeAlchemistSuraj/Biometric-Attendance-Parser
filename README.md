Biometric-Attendance-Parser

Overview

Biometric-Attendance-Parser is a Java-based application designed to process and consolidate biometric attendance data from multiple Excel sheets into a single, clean output file. The application handles Excel files containing employee attendance records, removes redundant header rows, and merges employee data blocks while ensuring a standardized "Days" row at the top, tailored to the specified month and year.

Features





Input Processing: Reads multiple Excel files (.xlsx) from a specified input directory.



Header Removal: Removes predefined header rows (e.g., "Monthly Status Report", company details, etc.) from each sheet.



Days Row Handling: Extracts or generates a "Days" row for the specified month and year, ensuring it appears only once at the top of the merged output.



Employee Data Consolidation: Stacks employee data blocks from all sheets into a single output sheet, preserving relevant attendance details.



Dynamic Month/Year Support: Prompts the user for the month and year to determine the number of days and generate the appropriate "Days" row if not present.



Apache POI Integration: Utilizes the Apache POI library for robust Excel file manipulation.

Prerequisites





Java: JDK 8 or higher.



Apache POI: Required for Excel file processing. Include the following dependencies in your project:

<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.2.3</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.3</version>
</dependency>



Input Directory: A directory named input containing the Excel files (.xlsx) to be processed.

Installation





Clone the repository:

git clone https://github.com/your-username/Biometric-Attendance-Parser.git



Navigate to the project directory:

cd Biometric-Attendance-Parser



Ensure the input directory exists and contains the Excel files to process.



Build the project using Maven (if using Maven):

mvn clean install

Usage





Place all input Excel files in the input directory.



Run the application:

java -jar Biometric-Attendance-Parser.jar



When prompted, enter the month (1-12) and year (e.g., 2025) for which the attendance data is being processed.



The application will:





Process all .xlsx files in the input directory.



Remove redundant headers and duplicate "Days" rows.



Merge employee data into a single output file named merged.xlsx in the project root directory.

Input Excel File Structure

Each input Excel file is expected to contain:





Optional header rows (e.g., "Monthly Status Report", "Company:", date ranges) that will be removed.



An optional "Days" row specifying the days of the month (e.g., "1 F", "2 St", etc.), which may vary based on the month’s length.



Employee data blocks containing details such as:





Department



Employee ID and name



Attendance status (e.g., "P" for Present, "WO" for Weekly Off)



InTime, OutTime, Duration, etc.

Example "Days" row for August 2025:

"Days","","1 F","2 St","3 S","4 M","5 T","6 W","7 Th","8 F","9 St","10 S",...

Output

The output is a single Excel file (merged.xlsx) with:





A single "Days" row at the top, either copied from the first sheet or generated based on the user-specified month and year.



All employee data blocks stacked vertically, with redundant headers and duplicate "Days" rows removed.

Example Run

Enter the month (1-12): 8
Enter the year (e.g., 2025): 2025
Merged file created: merged.xlsx

Limitations





Assumes all input Excel files are in .xlsx format and located in the input directory.



The "Days" row generation assumes a standard format (e.g., "1 F" for Friday, etc.) and may need adjustment for custom day abbreviations.



Handles only the first sheet of each Excel file.

Contributing

Contributions are welcome! Please submit a pull request or open an issue to discuss improvements or bug fixes.

License

This project is licensed under the MIT License. See the LICENSE file for details.
