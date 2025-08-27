import os
import pandas as pd

def insert_glass_types_multiple_outputs(inp_data, row_num):
    # If inp_data is a string, convert to list of lines
    if isinstance(inp_data, str):
        inp_data = inp_data.splitlines(keepends=True)

    # Define section markers
    start_marker = "Glass Types"
    end_marker = "Window Layers"
    start_marker1 = "Floors / Spaces / Walls / Windows / Doors"
    end_marker1 = "Electric & Fuel Meters"

    # Extract the climate from the file name (simulate path for climate extraction)
    inp_file_name = "example_Pune.inp"  # replace with actual filename logic if needed
    climate = inp_file_name.split('_')[-1].replace('.inp', '')  # e.g., 'Pune'

    # Load and filter the database
    database_path = 'database/AllData.xlsx'
    database = pd.read_excel(database_path, sheet_name='Glazing')
    filtered_db = database[database['Climate'] == climate]
    if filtered_db.empty:
        filtered_db = database[database['Climate'] == 'Others']

    # Validate row number
    if row_num >= len(filtered_db):
        print(f"Row number {row_num} out of range for climate '{climate}' in database.")
        return inp_data

    # Find Glass Types insertion section
    start_index, end_index = None, None
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 1
        if end_marker in line:
            end_index = i
            break
    if start_index is None or end_index is None:
        print("Glass Types section not found.")
        return inp_data

    # Find Windows section
    start_index1, end_index1 = None, None
    for i, line in enumerate(inp_data):
        if start_marker1 in line:
            start_index1 = i + 4
        if end_marker1 in line:
            end_index1 = i - 4
            break
    if start_index1 is None or end_index1 is None:
        print("Windows section not found.")
        return inp_data

    # Clean unwanted lines in WINDOW section
    updated_data1 = []
    in_window_section = False
    for i in range(start_index1, end_index1 + 1):
        line = inp_data[i].strip()
        if line.endswith("= WINDOW"):
            in_window_section = True
            updated_data1.append(inp_data[i])
            continue
        if in_window_section:
            if any(key in line for key in ["FRAME-WIDTH", "FRAME-CONDUCT", "SPACER-TYPE"]):
                continue
            if line == "..":
                in_window_section = False
        updated_data1.append(inp_data[i])
    inp_data = inp_data[:start_index1] + updated_data1 + inp_data[end_index1 + 1:]

    # Get selected row
    selected_row = filtered_db.iloc[row_num]

    # Prepare insertion content
    insertion_lines = [
        '$ ---------------------------------------------------------\n\n',
        f'"{selected_row["GLASS-TYPE"]}" = GLASS-TYPE\n',
        '   TYPE             = SHADING-COEF\n',
        f'   SHADING-COEF     = {round(selected_row["SHADING-COEF"],2)}\n',
        f'   GLASS-CONDUCT    = {round(selected_row["GLASS-CONDUCT"],2)}\n',
        f'   VIS-TRANS        = {selected_row["VIS-TRANS"]}\n',
        '   ..\n\n',
        '$ ---------------------------------------------------------\n'
    ]

    # Insert glass type lines
    updated_data = inp_data[:start_index] + insertion_lines + inp_data[end_index:]

    # Update GLASS-TYPE in WINDOW sections
    in_window_section = False
    for i in range(start_index1, end_index1 + 1):
        line = updated_data[i].strip()
        if line.endswith("= WINDOW"):
            in_window_section = True
        elif in_window_section and line.startswith("GLASS-TYPE"):
            updated_data[i] = f'   GLASS-TYPE       = "{selected_row["GLASS-TYPE"]}"\n'
            in_window_section = False

    return updated_data
