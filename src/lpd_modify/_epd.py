import streamlit as st

def perging_data_weekly(data, epd):
    # Define markers to identify the section of interest
    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    # Find the start and end indices of the relevant section
    start_index = None
    end_index = None
    for i, line in enumerate(data):
        if start_marker in line:
            start_index = i + 3
        if end_marker in line:
            end_index = i - 3
            break

    # Ensure markers were found
    if start_index is None or end_index is None:
        raise ValueError("Could not find the specified markers in the file.")

    # Iterate through the relevant section and replace `LIGHTING-W/AREA` values
    for i in range(start_index, end_index + 1):
        if "EQUIPMENT-W/AREA" in data[i]:
            # Replace the value of `LIGHTING-W/AREA` with the provided `lpd`
            data[i] = f"   EQUIPMENT-W/AREA  = ( {epd} )\n"
    
    # Return the modified data as a list of lines
    return data