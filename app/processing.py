import os
import pandas as pd
import xlsxwriter
import openpyxl
from openpyxl.styles import PatternFill

# Define paths for generated files
BASE_DIR = os.getcwd()  # Ensure correct working directory
INPUT_FILE = os.path.join(BASE_DIR, "slot_allocation.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "output_kavach_slots_colored.xlsx")

# Step 1: Slot Allocation Logic
def allocate_slots(stations, max_slots=45, max_frequencies=7):
    allocations = []
    current_frequency = 1
    station_alloc = [0] * max_slots
    onboard_alloc = [0] * max_slots

    def next_frequency():
        nonlocal current_frequency
        current_frequency += 1
        if current_frequency > max_frequencies:
            current_frequency = 1
        station_alloc[:] = [0] * max_slots
        onboard_alloc[:] = [0] * max_slots

    for station in stations:
        station_name = station["name"]
        station_slots = station["stationSlots"]
        onboard_slots = station["onboardSlots"]
        station_slot_range = []
        onboard_slot_allocations = []

        available_station_slots = station_alloc.count(0)
        available_onboard_slots = onboard_alloc.count(0)

        if station_slots > available_station_slots or onboard_slots > available_onboard_slots:
            next_frequency()

        allocated_station_slots = 0
        for i in range(max_slots):
            if station_alloc[i] == 0 and allocated_station_slots < station_slots:
                station_alloc[i] = station_name
                station_slot_range.append(f"P{i+1}")
                allocated_station_slots += 1

        allocated_onboard_slots = 0
        i = 0
        while allocated_onboard_slots < onboard_slots and i < max_slots:
            if station_alloc[i] == station_name:
                i += 1
                continue
            if onboard_alloc[i] == 0:
                onboard_alloc[i] = station_name
                onboard_slot_allocations.append(f"P{i+1}")
                allocated_onboard_slots += 1
            i += 1

        allocations.append({
            "Station": station_name,
            "Frequency": current_frequency,
            "Stationary Kavach Slots": ", ".join(station_slot_range),
            "Onboard Kavach Slots": ", ".join(onboard_slot_allocations)
        })

    return allocations  # This list will be used to generate Excel

# Step 2: Generate Excel from Allocated Data
def generate_excel(stations):
    allocations = allocate_slots(stations)  # Ensure slot allocation happens first

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

# Step 3: Apply Color Scheme to the Generated Excel
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
