
from flask import Flask, render_template, request, jsonify, send_file
import xlsxwriter
import os

app = Flask(__name__)

# Slot allocation logic
def allocate_slots(stations, max_slots=45, max_frequencies=7):
    allocations = []
    current_frequency = 1

    # Initialize allocation trackers for stationary and onboard Kavach slots
    station_alloc = [0] * max_slots
    onboard_alloc = [0] * max_slots

    def next_frequency():
        nonlocal current_frequency
        current_frequency += 1 
        if current_frequency > max_frequencies:
            current_frequency = 1
        # Reset slot allocations
        station_alloc[:] = [0] * max_slots
        onboard_alloc[:] = [0] * max_slots
        

    for station in stations:
        station_name= station["name"] 
        station_slots=station["stationSlots"]
        onboard_slots = station["onboardSlots"]
        station_slot_range = []  # To hold allocated stationary slot range for each station
        onboard_slot_allocations = []  # To hold allocated onboard slots for each station

        # Check if there are enough available slots, else move to next frequency
        available_station_slots = station_alloc.count(0)
        available_onboard_slots = onboard_alloc.count(0)

        if station_slots > available_station_slots or onboard_slots > available_onboard_slots:
            next_frequency()
            available_station_slots = station_alloc.count(0)
            available_onboard_slots = onboard_alloc.count(0)

        # Allocate stationary Kavach slots
        allocated_station_slots = 0
        for i in range(max_slots):
            if station_alloc[i] == 0 and allocated_station_slots < station_slots:
                station_alloc[i] = station_name
                station_slot_range.append(f"P{i+1}")
                allocated_station_slots += 1

        # Allocate onboard Kavach slots dynamically
        allocated_onboard_slots = 0
        available_onboard_slot_monitoring = available_onboard_slots
        onboard_slots_monitoring = onboard_slots

        i = 0
        while allocated_onboard_slots < onboard_slots and i < max_slots:
            # Skip if the slot is already occupied by stationary allocation
            if station_alloc[i] == station_name:
                i += 1
                continue

            if available_onboard_slot_monitoring / 2 >= onboard_slots:
                # Alternate Allocation
                if onboard_alloc[i] == 0:
                    onboard_alloc[i] = station_name
                    onboard_slot_allocations.append(f"P{i+1}")
                    allocated_onboard_slots += 1
                    # Skip the next slot to maintain alternation
                    available_onboard_slot_monitoring -= 1
                    onboard_slots_monitoring -= 1
                    i += 2
                else:
                    i += 1
            else:
                # Continuous Allocation
                if onboard_alloc[i] == 0:
                    onboard_alloc[i] = station_name
                    onboard_slot_allocations.append(f"P{i+1}")
                    allocated_onboard_slots += 1
                    available_onboard_slot_monitoring -= 1
                    onboard_slots_monitoring -= 1
                i += 1

        # Append station allocation details to the report
        allocations.append({
            "Station": station_name,
            "Frequency": current_frequency,
            "Stationary Kavach Slots": ", ".join(station_slot_range),
            "Onboard Kavach Slots": ", ".join(onboard_slot_allocations)
        })
    return allocations

@app.route("/index.html")
def home():
    return render_template("index.html")

@app.route("/allocate_slots_endpoint", methods=["POST"])
def allocate_slots_endpoint():
    try:
        stations = request.json
        allocations = allocate_slots(stations)

        # Create Excel file
        file_path = "slot_allocation.xlsx"
        workbook = xlsxwriter.Workbook(file_path)
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
        return jsonify({"fileUrl": f"/download/{file_path}"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download/<path:filename>")
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
