import os
import pandas as pd
import re

def updateEquipment(inp_data_str, equip_idx): 
    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    # --- Load EquipmentPower-Improvement database ---
    try:
        database = pd.read_excel("database/ML_ScaleUp_v02.xlsx", sheet_name='EquipmentPower-Improvement')
        percent = database['Percent'].iloc[equip_idx]
    except FileNotFoundError:
        raise FileNotFoundError("Excel file not found: database/ML_ScaleUp_v02.xlsx")
    except ValueError:
        raise ValueError("Sheet 'EquipmentPower-Improvement' not found in the Excel file.")
    except IndexError:
        raise IndexError(f"Invalid index provided: {equip_idx}. Total rows: {len(database)}")

    inp_data = inp_data_str.splitlines(keepends=True)

    if equip_idx == 0:
        return ''.join(inp_data)

    # --- Parse Global Parameters ---
    global_params = {}
    for line in inp_data:
        match = re.match(r'\s*"([^"]+)"\s*=\s*([0-9.]+)', line)
        if match:
            param_name = match.group(1).strip()
            param_value = float(match.group(2))
            global_params[param_name] = param_value

    # --- Locate SPACE Section ---
    start_index, end_index = None, None
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 4
        if end_marker in line and start_index is not None:
            end_index = i - 4
            break

    if start_index is None or end_index is None:
        raise ValueError("Could not find SPACE section markers in INP file.")

    modified_inp = inp_data.copy()

    for i in range(start_index, end_index):
        line = modified_inp[i]

        if "EQUIPMENT-W/AREA" in line:
            indent = line[:line.index("EQUIPMENT-W/AREA")]

            # Extract current EQUIPMENT-W/AREA value
            match = re.search(r'\(\s*([^)]*)\s*\)', line)
            if not match:
                continue
            cur_val_str = match.group(1).strip()

            # Remove braces if present
            cur_val_str = cur_val_str.strip("{} ").rstrip(")")

            cur_val = None

            # Case 1: single numeric
            try:
                cur_val = float(cur_val_str)
            except ValueError:
                # Case 2: list of numbers → take maximum
                if "," in cur_val_str:
                    try:
                        nums = [float(x.strip()) for x in cur_val_str.split(",") if x.strip()]
                        if nums:
                            cur_val = max(nums)
                    except ValueError:
                        pass  # fall through if parsing fails

                # Case 3: parameter reference
                if cur_val is None:
                    param_match = re.match(r'#?PA\("([^"]+)"', cur_val_str)
                    if param_match:
                        param_name = param_match.group(1).strip()
                        cur_val = global_params.get(param_name)
                        if cur_val is None:
                            raise ValueError(f"Parameter {param_name} not found in Global Parameters.")
                    else:
                        raise ValueError(f"Unexpected EQUIPMENT-W/AREA format: {cur_val_str}")

            # --- Apply formula ---
            new_val = cur_val * (1 - percent)

            # --- Replace EQUIPMENT-W/AREA ---
            new_equip_line = f"{indent}EQUIPMENT-W/AREA  = ( {new_val:.2f} )\n"
            modified_inp[i] = new_equip_line

            # --- Preserve AREA/PERSON if it exists ---
            area_person_match = re.search(r'(AREA/PERSON\s*=\s*\{?#?PA\(".*?"\)\}?)', line)
            if area_person_match:
                new_area_line = f"{indent}{area_person_match.group(1)}\n"
                modified_inp.insert(i + 1, new_area_line)

    return ''.join(modified_inp)