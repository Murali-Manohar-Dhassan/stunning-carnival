import os
import pandas as pd
import xlsxwriter
import openpyxl
from openpyxl.styles import PatternFill

# Define file paths
BASE_DIR = os.getcwd()
INPUT_FILE = os.path.join(BASE_DIR, "slot_allocation.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "output_kavach_slots_colored.xlsx")

# Step 1: Slot Allocation
def allocate_slots(stations, max_slots=45, max_frequencies=7):
    allocations = {station["name"]: [""] * max_slots for station in stations}  # Initialize slots for each station
    slot_numbers = [f"P{i}" for i in range(1, max_slots + 1)]
    current_frequency = 1
    station_alloc = [0] * max_slots  # Track allocated slots
    onboard_alloc = [0] * max_slots  # Track onboard slots

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

        available_station_slots = station_alloc.count(0)
        available_onboard_slots = onboard_alloc.count(0)

        if station_slots > available_station_slots or onboard_slots > available_onboard_slots:
            next_frequency()

        allocated_station_slots = 0
        for i in range(max_slots):
            if station_alloc[i] == 0 and allocated_station_slots < station_slots:
                station_alloc[i] = station_name
                allocations[station_name][i] = slot_numbers[i]
                allocated_station_slots += 1

        allocated_onboard_slots = 0
        for i in range(max_slots):
            if onboard_alloc[i] == 0 and allocated_onboard_slots < onboard_slots:
                onboard_alloc[i] = station_name
                allocations[station_name][i] = slot_numbers[i]
                allocated_onboard_slots += 1

    return allocations

# Step 2: Generate Excel File
def generate_excel(stations):
    allocations = allocate_slots(stations)
    df = pd.DataFrame(allocations)
    df.insert(0, "Slot", [f"P{i}" for i in range(1, len(df) + 1)])  # Add Slot column
    df.to_excel(INPUT_FILE, index=False)

    # Apply color scheme
    apply_color_scheme()
    
    return OUTPUT_FILE  # Return final colored file

# Step 3: Apply Color Coding
def apply_color_scheme():
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError("Generated Excel file not found!")

    df = pd.read_excel(INPUT_FILE)
    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb.active

    # Define color map for frequency
    color_map = {
        1: "FFFF00",  # Yellow
        2: "0000FF",  # Blue
        3: "FFA500",  # Orange
        4: "FF0000",  # Red
        5: "800080",  # Purple
        6: "FFC0CB",  # Pink
        7: "008000"   # Green
    }

    # Apply colors based on frequency values
    for row in range(2, len(df) + 2):  # Skip header row
        slot_name = ws.cell(row=row, column=1).value
        for col in range(2, ws.max_column + 1):  # Iterate through stations
            station_name = ws.cell(1, col).value
            if pd.notna(station_name) and pd.notna(slot_name):  # Check if cell contains a value
                # Extract frequency (assume frequency is stored in station data)
                frequency = 1  # Default frequency; modify as needed
                color_code = color_map.get(frequency, "FFFFFF")  # Default to white
                ws.cell(row=row, column=col).fill = PatternFill(start_color=color_code, end_color=color_code, fill_type="solid")

    # Save final file
    wb.save(OUTPUT_FILE)
