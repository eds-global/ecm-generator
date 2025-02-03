import streamlit as st

def getEPD(data, epd):
    if epd is None:
        return data
    
    with open(data, 'r') as file:
        data = file.read()
        
    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    start_index = data.find(start_marker)
    end_index = data.find(end_marker)

    if start_index == -1 or end_index == -1:
        raise ValueError("Could not find the specified markers in the file.")

    # Iterate through the relevant section and replace `LIGHTING-W/AREA` values
    for i in range(start_index, end_index + 1):
        if "EQUIPMENT-W/AREA" in data[i]:
            # Replace the value of `LIGHTING-W/AREA` with the provided `lpd`
            data[i] = f"   EQUIPMENT-W/AREA  = ( {epd} )\n"
    
    return data