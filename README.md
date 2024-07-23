### Excel Automation Organization Scripts For Flight Data

I used these scripts in my internship at Cebu Pacific in order to organize useful data and remove useless data

### Remove Blank Rows in Excel Sheets - "blankspaceremover.py"

This script uses the `openpyxl` library to remove completely blank rows from all worksheets in an Excel workbook.

#### Features:
- **Identifies Blank Rows:** Searches for rows that are completely blank in all columns.
- **Removes Blank Rows:** Deletes identified blank rows.
- **Works on All Sheets:** Applies the row removal process to all worksheets within the workbook.

This script is useful for cleaning up Excel files by removing unnecessary blank rows, ensuring the data remains concise and organized.

### Consolidate Data from Multiple Sheets into One - "consolidatedata.py"

This script uses the `openpyxl` library to consolidate data from multiple sheets within an Excel workbook into a single sheet. It ensures that there is no blank space between the consolidated data from each sheet and that only the first occurrence of the header is kept.

#### Features:
- **Consolidates Data:** Merges data from all sheets into a single sheet.
- **Removes Blank Spaces:** Ensures there are no blank spaces between data from different sheets.
- **Single Header:** Retains only the first occurrence of the header.

This script is useful for combining data from multiple sheets into a single sheet, ensuring a seamless and organized dataset without redundant headers or blank spaces.

### Convert Military Time to 12-Hour Format - "correctingmilitarytime.py"

This script uses the `openpyxl` library to convert date and time from military (24-hour) format to 12-hour format with AM/PM in an Excel file.

#### Features:
- **Converts Time Format:** Transforms date and time values in the second column from `YYYY-MM-DD HH:MM:SS` to `YYYY-MM-DD HH:MM:SS AM/PM`.
- **Handles Errors Gracefully:** Skips rows with incorrect date and time formats and provides informative error messages.
- **Saves Updated File:** Writes the converted time values to a new Excel file specified by the user.

This script is useful for reformatting date and time values in Excel files, making them easier to read and interpret by converting them to a 12-hour format.

### Reformat Date in Excel Sheets - "dateformatchange.py"

This script uses the `openpyxl` library to reformat dates in an Excel file, converting them to a specific format with an extra space between the date and time.

#### Features:
- **Reformats Dates:** Transforms date values in the first column from various known formats to `M/D/Y  H:MM:SS PM/AM`.
- **Handles Multiple Date Formats:** Supports parsing dates from several common formats to ensure compatibility.
- **Saves Updated File:** Writes the reformatted dates to a new Excel file specified by the user.

This script is useful for standardizing date formats in Excel files, making them consistent and easier to work with.

#### Notes
- Ensure the input Excel file has date values in the first column.
- The script will skip any rows where the date format does not match the expected formats and will print a message indicating the row number and the problematic value.

### Update Date with Time in Excel Sheets - "datetimecombine.py"

This script uses the `openpyxl` library to update the "DATE" column in an Excel file by combining the date from the "DATE" column with the time from the "TIME STARTED (UTC)" column.

#### Features:
- **Combines Date and Time:** Merges the date from the "DATE" column with the time from the "TIME STARTED (UTC)" column to create a new datetime value.
- **Handles Multiple Sheets:** Applies the date and time update process to all worksheets within the workbook.
- **Removes Incomplete Rows:** Deletes rows where the "TIME STARTED (UTC)" value is missing or the date cannot be parsed.

This script is useful for standardizing date and time data in Excel files, making it easier to work with combined datetime values.

### Compare and Merge Excel Sheets - "excelcomparer.py"

This script uses the `openpyxl` library to compare two Excel sheets and merge their data based on common columns. Specifically, it matches rows from the first sheet with rows from the second sheet based on "STATION" and "DATE" columns in the first sheet and "airport_code" and "time" columns in the second sheet.

#### Features:
- **Column Matching:** Ensures necessary columns exist in both sheets before proceeding.
- **Data Merging:** Adds headers from the second sheet to the first sheet and merges matching rows.
- **Efficient Lookup:** Uses a dictionary for fast row lookups from the second sheet.

This script is useful for combining data from two Excel files where specific columns act as keys for matching rows.

### Remove Rows Below "Departure" and "Arrival" in Excel Sheets - "findreplaceword.py"

This script uses the `openpyxl` library to identify and remove all rows below the first occurrence of the titles "DEPARTURE" or "ARRIVAL" in all worksheets of an Excel workbook.

#### Features:
- **Identifies Key Rows:** Searches for the row containing the titles "DEPARTURE" or "ARRIVAL".
- **Removes Subsequent Rows:** Deletes all rows below the identified row.
- **Works on All Sheets:** Applies the removal process to all worksheets within the workbook.

This script is useful for cleaning up Excel files by removing unnecessary rows below critical headers such as "DEPARTURE" and "ARRIVAL", ensuring the data remains relevant and organized.

### Reformat Date to Custom Format in Excel Sheets - "reformat_date_to_custom_format.py"

This script uses the `openpyxl` library to reformat dates in the "DATE" column of an Excel file to a custom format (`'%Y-%m-%dT%H:%M:%SZ'`). It processes all worksheets within the workbook and handles various date formats.

#### Features:
- **Date Reformatting:** Converts date values to the custom format `'%Y-%m-%dT%H:%M:%SZ'`.
- **Handles Multiple Formats:** Parses and reformats dates from multiple common formats.
- **Sheet-Wide Application:** Applies the reformatting to all worksheets within the workbook.
- **Row Deletion:** Identifies and deletes rows with unparsable date values to ensure data integrity.

This script is useful for standardizing date formats in Excel files to a specific, custom format that is consistent and easily usable in various applications.

### Remove Rows with Empty Columns in Excel Sheets - "removeblank.py"

This script uses the `openpyxl` library to remove rows from all worksheets in an Excel workbook where any of the first three columns are empty.

#### Features:
- **Identifies Rows with Empty Columns:** Searches for rows where any of the first three columns contain empty cells.
- **Removes Identified Rows:** Deletes rows that meet the criteria.
- **Works on All Sheets:** Applies the row removal process to all worksheets within the workbook.

This script is useful for cleaning up Excel files by removing rows with incomplete data in the first three columns, ensuring the dataset is complete and organized.

### Remove Leading Zeros from Date-Time Column in Excel Sheets - "removezeroes.py"

This script uses the `openpyxl` library to remove leading zeros from date-time values in the second column of an Excel file. It processes all rows in the active worksheet and updates the date-time format accordingly.

#### Features:
- **Leading Zero Removal:** Removes leading zeros from the month, day, and hour in date-time strings.
- **Error Handling:** Skips rows with unparsable date-time values and provides informative messages.

This script is useful for cleaning up date-time values in Excel files by removing unnecessary leading zeros, ensuring a consistent and more readable format.

### Remove Timezone from Time Column in Excel Sheets - "replaceparts.py"

This script uses the `openpyxl` library to remove the timezone component from the "time" column in an Excel file. The script processes all worksheets within the workbook, updating date-time strings by removing the 'T' and 'Z' characters.

#### Features:
- **Timezone Removal:** Removes 'T' and 'Z' characters from date-time strings in the "time" column.
- **Sheet-Wide Application:** Applies the changes to all worksheets within the workbook.
- **Error Handling:** Skips sheets that do not contain a "time" column and provides informative messages.

This script is useful for cleaning up date-time strings in Excel files by removing unnecessary timezone information, making the data easier to work with and more readable.

### Adjust DATE and STATION Columns in Excel Sheets - "swapcolumns.py"

This script uses the `openpyxl` library to ensure that the "DATE" and "STATION" columns are positioned correctly in all worksheets of an Excel workbook. Specifically, it ensures that "DATE" is in column A and "STATION" is in column B, swapping them if necessary.

#### Features:
- **Identifies Columns:** Searches for the columns containing "DATE" and "STATION" headers.
- **Swaps Columns:** Ensures "DATE" is in column A and "STATION" is in column B, swapping the columns if needed.
- **Works on All Sheets:** Applies the adjustments to all worksheets within the workbook.

This script is useful for ensuring that critical columns like "DATE" and "STATION" are correctly positioned in your Excel sheets, facilitating better data organization and consistency across all worksheets.

### Reset Time to Midnight in Excel Sheets - "turnhourminutetimetohour.py"

This script uses the `openpyxl` library to reset the time component of dates in the "DATE" column of an Excel file to midnight (`00:00:00`). The script processes all rows in the active worksheet and updates the date-time format accordingly.

#### Features:
- **Time Reset:** Sets the time component of date-time values to midnight (`00:00:00`).
- **Handles Different Formats:** Parses date-time values in a specific format and updates them.
- **Error Handling:** Skips rows with unparsable date-time values and provides informative messages.

This script is useful for standardizing the time component of date-time values to midnight in Excel files, ensuring a consistent format for further data processing or analysis.

### Unmerge Cells in Excel Sheets - "unmergedcells.py"

This script uses the `openpyxl` library to unmerge cells in all worksheets of an Excel workbook. When cells are unmerged, the values from the merged cells are consolidated into a single string, which is then placed in each of the previously merged cells.

#### Features:
- **Unmerges Cells:** Automatically identifies and unmerges all merged cells in the workbook.
- **Consolidates Values:** Combines the values of merged cells into a single string and places it in each of the previously merged cells.
- **Works on All Sheets:** Applies the unmerging process to all worksheets within the workbook.

This script is useful for cleaning up Excel files by unmerging cells and ensuring that all data remains visible and accessible in each previously merged cell.
