import os
import pandas as pd
import streamlit as st

def insert_glass_types_multiple_outputs(inp_file_path):
    # Define the markers
    start_marker = "Glass Types"
    end_marker = "Window Layers"

    start_marker1 = "Floors / Spaces / Walls / Windows / Doors"
    end_marker1 = "Electric & Fuel Meters"

    start_marker2 = "Floors / Spaces / Walls / Windows / Doors"
    end_marker2 = "Electric & Fuel Meters"
    
    # Extract the climate from the input file name
    inp_file_name = os.path.basename(inp_file_path)
    inp_file_base_name = inp_file_name.replace('.inp', '')  # Remove `.inp` for output naming
    climate = inp_file_name.split('_')[-1].replace('.inp', '')  # Extract the last word before `.inp`
    
    # Load the database
    database = pd.read_excel('database/ML_ScaleUp_v02.xlsx', sheet_name="Glazing-Options")
    filtered_db = database[database['Climate'] == climate]
    
    # If no match is found, default to 'Others'
    if filtered_db.empty:
        filtered_db = database[database['Climate'] == 'Others']
    
    # Read the .inp file
    with open(inp_file_path, 'r') as file:
        inp_data = file.readlines()
    
    # Find the range of insertion
    start_index, end_index = None, None
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 1
        if end_marker in line:
            end_index = i
            break
    
    if start_index is None or end_index is None:
        print(f"Error: Start or end marker not found in the .inp file: {inp_file_path}")
        return

    start_index1, end_index1 = None, None
    for i, line in enumerate(inp_data):
        if start_marker1 in line:
            start_index1 = i + 4  # Start index is 4 lines below the start marker
        if end_marker1 in line:
            end_index1 = i - 4  # End index is 4 lines above the end marker
            break

    if start_index1 is None or end_index1 is None:
        print("Error: Start or end marker not found for Windows section.")
        return inp_data

    start_index2, end_index2 = None, None
    for i, line in enumerate(inp_data):
        if start_marker2 in line:
            start_index2 = i + 4  # Start index is 4 lines below the start marker
        if end_marker2 in line:
            end_index2 = i - 4  # End index is 4 lines above the end marker
            break

    if start_index2 is None or end_index2 is None:
        print("Error: Start or end marker not found for Windows section.")
        return inp_data

    updated_data1 = []
    in_window_section = False
    for i in range(start_index2, end_index2 + 1):
        line = inp_data[i].strip()

        # Detect the start of a WINDOW section
        if line.endswith("= WINDOW"):
            in_window_section = True
            updated_data1.append(inp_data[i])  # Keep the WINDOW header
            continue

        # If in WINDOW section, exclude specified lines
        if in_window_section:
            if any(key in line for key in ["FRAME-WIDTH", "FRAME-CONDUCT", "SPACER-TYPE"]):
                continue  # Skip the line
            if line == "..":
                in_window_section = False

        updated_data1.append(inp_data[i])
    inp_data = inp_data[:start_index2] + updated_data1 + inp_data[end_index2 + 1:]
    
    # Create multiple output files for each row in the filtered database
    for idx, (_, row) in enumerate(filtered_db.iterrows(), start=1):
        # Prepare the insertion content
        insertion_lines = [
            '$ ---------------------------------------------------------\n\n',
            f'"{row["GLASS-TYPE"]}" = GLASS-TYPE\n',
            f'   TYPE             = SHADING-COEF\n',
            f'   SHADING-COEF     = {row["SHADING-COEF"]}\n',
            f'   GLASS-CONDUCT    = {row["GLASS-CONDUCT"]}\n',
            f'   VIS-TRANS        = {row["VIS-TRANS"]}\n',
            '   ..\n\n',
            '$ ---------------------------------------------------------\n'
        ]
        
        # Insert the new content
        updated_data = inp_data[:start_index] + insertion_lines + inp_data[end_index:]

        # Modify the = WINDOW sections
        in_window_section = False
        for i in range(start_index1, end_index1 + 1):
            line = updated_data[i].strip()
            if line.endswith("= WINDOW"):
                in_window_section = True
            elif in_window_section:
                # Replace GLASS-TYPE in the window section
                if line.startswith("GLASS-TYPE"):
                    updated_data[i] = f'   GLASS-TYPE       = "{row["GLASS-TYPE"]}"\n'
                    in_window_section = False  # Stop after replacing GLASS-TYPE
 
        with open(inp_file_path, 'w') as file:
            file.write(updated_data)
    
    return inp_file_path

def processFolder(file_name, window):
    if file_name.endswith(".inp"):
        with open(file_name, 'r') as file:
            inp_data = file.readlines()
        inp_data = insert_glass_types_multiple_outputs(inp_data)
        
        return inp_data

def getWindows(inp_path, window):
    if window is None:
        return inp_path
    


    processFolder(inp_path, window)