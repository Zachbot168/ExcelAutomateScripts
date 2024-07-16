### Excel Automation Organization Scripts For Flight Data

I used these scripts in my internship at Cebu Pacific in order to organize useful data and remove useless data

### Unmerge Cells in Excel Sheets - "unmergedcells.py"

This script uses the `openpyxl` library to unmerge cells in all worksheets of an Excel workbook. When cells are unmerged, the values from the merged cells are consolidated into a single string, which is then placed in each of the previously merged cells.

#### Features:
- **Unmerges Cells:** Automatically identifies and unmerges all merged cells in the workbook.
- **Consolidates Values:** Combines the values of merged cells into a single string and places it in each of the previously merged cells.
- **Works on All Sheets:** Applies the unmerging process to all worksheets within the workbook.

This script is useful for cleaning up Excel files by unmerging cells and ensuring that all data remains visible and accessible in each previously merged cell.

### Remove Rows Below "Departure" and "Arrival" in Excel Sheets - "findreplaceword.py"

This script uses the `openpyxl` library to identify and remove all rows below the first occurrence of the titles "DEPARTURE" or "ARRIVAL" in all worksheets of an Excel workbook.

#### Features:
- **Identifies Key Rows:** Searches for the row containing the titles "DEPARTURE" or "ARRIVAL".
- **Removes Subsequent Rows:** Deletes all rows below the identified row.
- **Works on All Sheets:** Applies the removal process to all worksheets within the workbook.

This script is useful for cleaning up Excel files by removing unnecessary rows below critical headers such as "DEPARTURE" and "ARRIVAL", ensuring the data remains relevant and organized.

### Adjust DATE and STATION Columns in Excel Sheets - "swapcolumns.py"

This script uses the `openpyxl` library to ensure that the "DATE" and "STATION" columns are positioned correctly in all worksheets of an Excel workbook. Specifically, it ensures that "DATE" is in column A and "STATION" is in column B, swapping them if necessary.

#### Features:
- **Identifies Columns:** Searches for the columns containing "DATE" and "STATION" headers.
- **Swaps Columns:** Ensures "DATE" is in column A and "STATION" is in column B, swapping the columns if needed.
- **Works on All Sheets:** Applies the adjustments to all worksheets within the workbook.

This script is useful for ensuring that critical columns like "DATE" and "STATION" are correctly positioned in your Excel sheets, facilitating better data organization and consistency across all worksheets.
