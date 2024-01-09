# Duplicate-File-Finder-and-Reporter

# Objective:
Identify duplicate files within a specified directory, generate an Excel report, and highlight duplicate file pairs with color-coded formatting.
# Code Overview:
The Python script utilizes hashlib and openpyxl libraries to calculate file hashes and create an Excel report highlighting duplicate files.
# Functionalities:
1. Hash Calculation:
    * Utilizes the MD5 hash algorithm to calculate file hashes.
2. Duplicate File Detection:
    * Recursively scans a specified directory for duplicate files based on their hashes.
    * Groups duplicate files into pairs.
3. Excel Report Generation:
    * Creates an Excel workbook with a worksheet named "Duplicate Files."
    * Highlights duplicate file pairs with alternating light colors.
    * Includes columns for file name, hash, and pair ID.
4. Output:
    * Saves the Excel report to the specified output folder with a user-defined filename.
# Usage:
1. Set the directory_path variable to the desired input folder path.
2. Set the output_path variable to the desired output folder path.
3. Customize the excel_filename variable for the Excel report.
# Note:
1. Ensure necessary Python libraries (hashlib, openpyxl) are installed.
2. Adjustments may be needed based on the file types and sizes in the specified directory.
3. The generated Excel report provides a visual representation of duplicate file pairs.
#  Example Output:
An Excel report named "DUP Bank (45S)_1.xlsx" is saved to the specified output folder, highlighting duplicate file pairs for easy identification.
# Output Location:
The generated Excel report is saved to the specified output folder, and the full path is printed to the console.
