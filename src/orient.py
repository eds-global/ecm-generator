import pandas as pd
import os
import sys
import streamlit as st

def resource_path(relative_path):
    """Get absolute path to resource (works in dev and exe)"""
    if getattr(sys, 'frozen', False):  # Running inside exe
        base_path = sys._MEIPASS
    else:
        # Go up one folder from src/ → project root
        base_path = os.path.dirname(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

def updateOrientation(data_lines, orient):
    # with open(inp_file_path, 'r') as file:
    #     data_lines = file.readlines()
    
    start_marker = '"Building Data" = BUILD-PARAMETERS'
    # end_marker = '$ ---------------------------------------------------------\n$              Materials / Layers / Constructions'
    end_marker = '$ --'

    # Find the index where "Building Data" starts
    building_start_index = None
    for i, line in enumerate(data_lines):
        if start_marker in line:
            building_start_index = i
            break

    if building_start_index is None:
        print("BUILDING DATA section not found. Skipping...")
        return None

    # Check if AZIMUTH is already present in the section
    azimuth_found = False
    insert_index = building_start_index + 1
    for i in range(building_start_index + 1, len(data_lines)):
        if "AZIMUTH" in data_lines[i]:
            azimuth_found = True
            azimuth_line_index = i
            break
        if end_marker in data_lines[i]:
            break

    # Read database and get new azimuth
    # database = pd.read_excel("database/ML_ScaleUp_v02.xlsx", sheet_name='Orientation')
    excel_path = resource_path(os.path.join("database", "ML_ScaleUp_v02.xlsx"))
    # st.write(excel_path)
    database = pd.read_excel(excel_path, sheet_name="Orientation")
    if orient >= len(database):
        print(f"Orientation index {orient} out of bounds. Skipping...")
        return None
    new_azimuth = database.loc[orient, 'Rotate']

    # Modify or insert AZIMUTH line
    new_data_lines = data_lines.copy()
    azimuth_line = f"   AZIMUTH          = {new_azimuth}\n"

    if azimuth_found:
        new_data_lines[azimuth_line_index] = azimuth_line
    else:
        new_data_lines.insert(insert_index, azimuth_line)

    return new_data_lines