import os
import time
import pandas as pd
from app.frequencyAllotment import allocate_slots
from app.ColorExcel import apply_frequency_coloring
from app.colorCodingScheme import apply_slot_coloring

# Define file paths
BASE_DIR = os.getcwd()
INPUT_FILE = os.path.join(BASE_DIR, "slot_allocation.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "output_kavach_slots_colored.xlsx")

# Generate Excel from Slot Allocation
def generate_excel(stations):
    allocations = allocate_slots(stations)  # Use your existing function
    
    # Convert allocations into DataFrame
    df = pd.DataFrame(allocations)
    
    # Save allocation to Excel
    df.to_excel(INPUT_FILE, index=False)

    # Ensure file is fully written before proceeding
    wait_for_file(INPUT_FILE)

    # Apply color formatting using existing scripts
# Ensure the file is fully written before applying coloring
max_wait_time = 10
wait_time = 0
while not os.path.exists(INPUT_FILE) and wait_time < max_wait_time:
    time.sleep(1)
    wait_time += 1

if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(f"Expected file '{INPUT_FILE}' not found. Ensure slot allocation runs first.")

apply_frequency_coloring(INPUT_FILE)
apply_slot_coloring(INPUT_FILE)


    return OUTPUT_FILE  # Return final colored Excel file

# Ensure the file is created before proceeding
def wait_for_file(file_path, max_wait_time=10):
    wait_time = 0
    while not os.path.exists(file_path) and wait_time < max_wait_time:
        time.sleep(1)
        wait_time += 1

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Expected file '{file_path}' not found after waiting {max_wait_time} seconds.")
