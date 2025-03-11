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

def generate_excel(stations):
    allocations = allocate_slots(stations)  # Call slot allocation logic
    
    # Convert allocations into DataFrame
    df = pd.DataFrame(allocations)
    
    # Save allocation to Excel
    df.to_excel(INPUT_FILE, index=False)

    # ✅ Log when the file is created
    print(f"✅ File Created: {INPUT_FILE}")
    
    # Ensure file is fully written before proceeding
    wait_for_file(INPUT_FILE)

    # Apply color formatting using existing scripts
    apply_frequency_coloring(INPUT_FILE)  # Use ColorExcel.py logic
    apply_slot_coloring(INPUT_FILE)  # Use colorCodingScheme.py logic

    return OUTPUT_FILE  # Return final colored Excel file

# Ensure the file is created before proceeding
def wait_for_file(file_path, max_wait_time=15):
    wait_time = 0
    while not os.path.exists(file_path) and wait_time < max_wait_time:
        print(f"⏳ Waiting for file: {file_path} ({wait_time}s elapsed)")
        time.sleep(1)
        wait_time += 1

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"❌ Expected file '{file_path}' not found after waiting {max_wait_time} seconds.")
