import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import os

BASE_DIR = os.getcwd()  # Get the current working directory

# Define a color map based on frequency values
color_map = {
    1: "FFFF00",  # Yellow
    2: "0000FF",  # Blue
    3: "FFA500",  # Orange
    4: "FF0000",  # Red
    5: "800080",  # Purple
    6: "FFC0CB",  # Pink
    7: "008000"   # Green
}

# Read the input data from the Excel file
input_file_path = os.path.join(BASE_DIR, "slot_allocation.xlsx")
import os
import time

# Wait for slot_allocation.xlsx to be created (max wait time: 10 seconds)
max_wait_time = 10
wait_time = 0
while not os.path.exists(input_file_path) and wait_time < max_wait_time:
    time.sleep(1)
    wait_time += 1

# Check again before reading
if not os.path.exists(input_file_path):
    raise FileNotFoundError(f"Expected file '{input_file_path}' not found after waiting {max_wait_time} seconds.")

input_df = pd.read_excel(input_file_path)

# Generate all possible slot names (P1 to P45)
all_slots = [f"P{i}" for i in range(1, 46)]

# Extract all unique station names
all_stations = input_df["Station"].unique()

# Initialize the output DataFrame with all slots and station columns
output_df = pd.DataFrame(index=all_slots, columns=["Slot"])
output_df["Slot"] = all_slots  # Populate the 'Slot' column with P1 to P45

# Populate the slots for each station
for _, row in input_df.iterrows():
    station = row["Station"]
    frequency = row["Frequency"]
    stationary_slots = str(row["Stationary Kavach Slots"]).split(", ")
    onboard_slots = str(row["Onboard Kavach Slots"]).split(", ")
    
    # Get the color based on the frequency
    color_code = color_map.get(frequency, "FFFFFF")  # Default to white if frequency not in range
    
    # Combine Stationary and Onboard slots into one column
    for slot in stationary_slots:
        if slot in all_slots:
            output_df.loc[slot, f"{station}"] = slot

    for slot in onboard_slots:
        if slot in all_slots:
            # Only populate if no entry exists (to avoid overwriting stationary slots)
            if pd.isna(output_df.loc[slot, f"{station}"]):
                output_df.loc[slot, f"{station}"] = slot

# Replace NaN with empty strings for better presentation
output_df.fillna("", inplace=True)

# Save the DataFrame to an Excel file using openpyxl to apply formatting
output_file_path = os.path.join(BASE_DIR, "output_kavach_slots_colored.xlsx")
output_df.to_excel(output_file_path, index=False)

# Load the saved workbook to apply color formatting
wb = openpyxl.load_workbook(output_file_path)
ws = wb.active

# Apply the color formatting for stationary slots based on frequency
for row in range(2, len(all_slots) + 2):  # Skip the header row
    for station in all_stations:
        # Check if the slot is a stationary slot for the current station
        cell_value = ws.cell(row=row, column=1).value
        if cell_value in str(input_df[input_df["Station"] == station]["Stationary Kavach Slots"].values[0]).split(", "):
            # Apply the color formatting for stationary slots
            frequency = input_df[input_df["Station"] == station]["Frequency"].values[0]
            color_code = color_map.get(frequency, "FFFFFF")
            # Get the corresponding column for the current station
            column_index = output_df.columns.get_loc(f"{station}") + 1
            cell = ws.cell(row=row, column=column_index)
            cell.fill = PatternFill(start_color=color_code, end_color=color_code, fill_type="solid")

# Save the workbook with color formatting applied
wb.save(output_file_path)

print(f"Output saved to {output_file_path}")
