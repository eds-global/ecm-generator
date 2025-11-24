import os
import pandas as pd
import re

def updateLPD(inp_data, light): 
    # --- Turn OFF DAYLIGHTING if YES ---
    for i, line in enumerate(inp_data):
        if "DAYLIGHTING" in line and "= YES" in line:
            inp_data[i] = re.sub(r'=\s*YES', '= NO', line)

    if light == 0:
        return ''.join(inp_data)

    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "$ --"

    try:
        database = pd.read_excel("database/ML_ScaleUp_v02.xlsx", sheet_name='LightingPower-Improvement')
        percent = database['Percent'].iloc[light]
    except FileNotFoundError:
        raise FileNotFoundError("Excel file not found: database/ML_ScaleUp_v02.xlsx")
    except ValueError:
        raise ValueError("Sheet 'LightingPower-Improvement' not found in the Excel file.")
    except IndexError:
        raise IndexError(f"Invalid 'light' index provided: {light}. Total rows: {len(database)}")

    # --- Parse Global Parameters ---
    global_params = {}
    for line in inp_data:
        match = re.match(r'\s*"([^"]+)"\s*=\s*([0-9.]+)', line)
        if match:
            param_name = match.group(1).strip()
            param_value = float(match.group(2))
            global_params[param_name] = param_value

    # --- Preprocess: Merge lines where LIGHTING-W/AREA and next line start with {, #, or PA ---
    merged_data = []
    i = 0
    while i < len(inp_data):
        line = inp_data[i]
        if "LIGHTING-W/AREA" in line and i + 1 < len(inp_data):
            next_line = inp_data[i + 1].lstrip()
            if next_line.startswith(("{", "#", "PA")):
                # Merge next line into current line
                line = line.rstrip("\n") + " " + next_line
                i += 1  # Skip the next line
        merged_data.append(line)
        i += 1
    inp_data = merged_data

    # --- Locate SPACE Section ---
    start_index, end_index = None, None
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 4
            break

    for i, line in enumerate(inp_data):
        if end_marker in line and start_index is not None and i > start_index:
            end_index = i - 4
            break
    
    if start_index is None or end_index is None:
        raise ValueError("Could not find SPACE section markers in INP file.")

    # --- Modify LIGHTING-W/AREA values ---
    modified_inp = inp_data.copy()
    for i in range(start_index, end_index):
        if "LIGHTING-W/AREA" in modified_inp[i]:
            line = modified_inp[i]
            
            # Extract value inside parentheses
            match = re.search(r'\(\s*([^)]*)\s*\)', line)
            if not match:
                continue
            cur_val_str = match.group(1).strip()

            cur_val_str = cur_val_str.strip("{} ").rstrip(")")

            # Case 1: direct numeric value
            try:
                cur_val = float(cur_val_str)
            except ValueError:
                # Case 2: parameter reference
                param_match = re.match(r'#?PA\("([^"]+)"', cur_val_str)
                if param_match:
                    param_name = param_match.group(1).strip()
                    cur_val = global_params.get(param_name)
                    if cur_val is None:
                        raise ValueError(f"Parameter {param_name} not found in Global Parameters.")
                else:
                    raise ValueError(f"Unexpected value format for LIGHTING-W/AREA: {cur_val_str}")
            
            new_val = cur_val * (1 - percent)
            modified_inp[i] = f"   LIGHTING-W/AREA  = ( {new_val:.2f} )\n"

    return ''.join(modified_inp)
