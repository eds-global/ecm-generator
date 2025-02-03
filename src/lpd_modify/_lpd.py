import streamlit as st

def getLPD(data, lpd):
    if lpd is None:
        return data
    
    with open(data, 'r') as file:
        data = file.read()

    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    start_index = data.find(start_marker)
    end_index = data.find(end_marker)

    if start_index == -1 or end_index == -1:
        raise ValueError("Could not find the specified markers in the file.")

    for i in range(start_index, end_index + 1):
        if "LIGHTING-W/AREA" in data[i]:
            data[i] = f"   LIGHTING-W/AREA  = ( {lpd} )\n"

    return data