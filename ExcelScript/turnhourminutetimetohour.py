import openpyxl
from datetime import datetime

def reset_time_to_midnight(file_path, output_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    # Assuming "DATE" is in the first column
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1):
        date_cell = row[0]  # This refers to the cell in the "DATE" column for each row

        if date_cell.value:
            try:
                # Check if the cell value is already a datetime object
                if isinstance(date_cell.value, datetime):
                    time_value = date_cell.value
                else:
                    # Parse the date-time value using the existing format
                    time_value = datetime.strptime(date_cell.value, "%m/%d/%Y  %I:%M:%S %p")

                # Reset the time to 00:00:00
                reset_time = time_value.replace(minute=0, second=0)

                # Format it to the desired output format
                formatted_time = reset_time.strftime("%-m/%-d/%Y  %-I:%M:%S %p")

                # Update the cell with the new format
                date_cell.value = formatted_time

            except ValueError:
                print(f"Skipping row {date_cell.row} due to format error: {date_cell.value}")

    wb.save(output_path)
    print(f"Updated times to midnight and saved to {output_path}")

# User input
file_path = input("Please enter the path to the Excel file: ")
output_path = input("Please enter the path to save the output Excel file: ")

reset_time_to_midnight(file_path, output_path)
