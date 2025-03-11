import os
import pandas as pd
import openpyxl
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

    # Apply color formatting using existing scripts
    apply_frequency_coloring(INPUT_FILE)  # Use ColorExcel.py logic
    apply_slot_coloring(INPUT_FILE)  # Use colorCodingScheme.py logic

    return OUTPUT_FILE  # Return final colored Excel file
