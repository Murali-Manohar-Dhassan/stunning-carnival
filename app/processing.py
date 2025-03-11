import os
import pandas as pd
import xlsxwriter
import openpyxl
from openpyxl.styles import PatternFill

# Define paths for generated files
BASE_DIR = os.getcwd()  # Ensure the app works in Render
INPUT_FILE = os.path.join(BASE_DIR, "slot_allocation.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "output_kavach_slots_colored.xlsx")

# First Step: Generate Excel from User Input
def generate_excel(allocations):
    workbook = xlsxwriter.Workbook(INPUT_FILE)
    worksheet = workbook.add_worksheet()

    headers = ["Station", "Frequency", "Stationary Kavach Slots", "Onboard Kavach Slots"]
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    for row, alloc in enumerate(allocations, start=1):
        worksheet.write(row, 0, alloc["Station"])
        worksheet.write(row, 1, alloc["Frequency"])
        worksheet.write(row, 2, alloc["Stationary Kavach Slots"])
        worksheet.write(row, 3, alloc["Onboard Kavach Slots"])

    workbook.close()

    # After generating the first file, apply the color scheme
    apply_color_scheme()
    
    return OUTPUT_FILE  # Return the final colored Excel file

# Second Step: Apply Color Scheme to the Generated Excel
def apply_color_scheme():
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError("Generated Excel file not found!")

    input_df = pd.read_excel(INPUT_FILE)

    # Define color mapping based on frequency
    color_map = {
        1: "FFFF00",  # Yellow
        2: "0000FF",  # Blue
        3: "FFA500",  # Orange
        4: "FF0000",  # Red
        5: "800080",  # Purple
        6: "FFC0CB",  # Pink
        7: "008000"   # Green
    }

    # Load Excel file to apply colors
    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb.active

    # Apply color formatting based on frequency
    for row in range(2, ws.max_row + 1):  # Skip header row
        station_name = ws.cell(row=row, column=1).value
        frequency = input_df[input_df["Station"] == station_name]["Frequency"].values[0]
        color_code = color_map.get(frequency, "FFFFFF")  # Default white if missing

        for col in range(2, ws.max_column + 1):  # Apply color to all station slots
            ws.cell(row=row, column=col).fill = PatternFill(start_color=color_code, end_color=color_code, fill_type="solid")

    # Save final colored Excel file
    wb.save(OUTPUT_FILE)
