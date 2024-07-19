import openpyxl
from datetime import datetime

def remove_leading_zeros(file_path, output_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2, min_col=2, max_col=2):  # Focus on the second column
        time_cell = row[0]  # This refers to the cell in the second column for each row

        if time_cell.value:
            try:
                # Parse the date-time value using the existing format
                time_value = datetime.strptime(time_cell.value, "%m/%d/%Y  %I:%M:%S %p")

                # Remove leading zeroes from month, day, and hour
                formatted_time = time_value.strftime("%-m/%-d/%Y  %-I:%M:%S %p")

                # Update the cell with the new format
                time_cell.value = formatted_time

            except ValueError:
                print(f"Skipping row {time_cell.row} due to format error: {time_cell.value}")

    wb.save(output_path)
    print(f"Removed leading zeroes and saved to {output_path}")

# User input
file_path = input("Please enter the path to the Excel file: ")
output_path = input("Please enter the path to save the output Excel file: ")

remove_leading_zeros(file_path, output_path)
