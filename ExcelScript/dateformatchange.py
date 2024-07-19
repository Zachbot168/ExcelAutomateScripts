import openpyxl
from datetime import datetime

def reformat_date(file_path, output_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2, max_col=1):  # Adjust max_col to the column number of your DATE column
        date_cell = row[1]  # Assuming "DATE" is in the first column

        if date_cell.value:
            try:
                # Try to parse the date using known formats
                try:
                    date_value = datetime.strptime(date_cell.value, "%Y-%m-%d %H:%M:%S %p")
                except ValueError:
                    try:
                        date_value = datetime.strptime(date_cell.value, "%Y-%m-%d %I:%M:%S %p")
                    except ValueError:
                        date_value = datetime.strptime(date_cell.value, "%Y-%m-%d %H:%M:%S")

                # Format it to M/D/Y H:MM:SS PM/AM with an extra space in the middle
                formatted_date = date_value.strftime("%m/%d/%Y  %I:%M:%S %p")

                # Update the cell with the new format
                date_cell.value = formatted_date

            except ValueError:
                print(f"Skipping row {row[0].row} due to format error: {date_cell.value}")

    wb.save(output_path)
    print(f"Converted dates to the new format and saved to {output_path}")

# User input
file_path = input("Please enter the path to the Excel file: ")
output_path = input("Please enter the path to save the output Excel file: ")

reformat_date(file_path, output_path)
