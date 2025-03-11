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

    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb.active
    df = pd.read_excel(INPUT_FILE)


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

    # Iterate through each column (stations) to apply colors
    for col in range(2, ws.max_column + 1):  # Start from column 2 (ignore 'Slot' column)
        station_name = ws.cell(row=1, column=col).value  # Get station name
        if not station_name:
            continue  # Skip if no station name

        # Get frequency of this station from the dataframe
        try:
            frequency = df[df["Slot"] == station_name]["Frequency"].values[0]
        except IndexError:
            frequency = 1  # Default to frequency 1 if not found

        # Determine the fill color based on frequency
        color_code = color_map.get(frequency, "FFFFFF")  # Default to white

        # Apply color only to non-empty cells
        for row in range(2, ws.max_row + 1):  # Skip header row
            cell = ws.cell(row=row, column=col)
            if cell.value:  # Only color non-empty cells
                cell.fill = PatternFill(start_color=color_code, end_color=color_code, fill_type="solid")
    # Save final file
    wb.save(OUTPUT_FILE)
