import os
import pandas as pd
import re

def updateEquipment(inp_data_str, light): 
    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    try:
        database = pd.read_excel("database/ML_ScaleUp_v02.xlsx", sheet_name='EquipmentPower-Improvement')
        improvement = database['Improvement'].iloc[light]
    except FileNotFoundError:
        raise FileNotFoundError("Excel file not found: database/ML_ScaleUp_v02.xlsx")
    except ValueError:
        raise ValueError("Sheet 'EquipmentPower-Improvement' not found in the Excel file.")
    except IndexError:
        raise IndexError(f"Invalid 'light' index provided: {light}. Total improvements: {len(database)}")

    inp_data = inp_data_str.splitlines(keepends=True)
    start_index, end_index = None, None

    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 4
        if end_marker in line:
            end_index = i - 4
            break

    if start_index is None or end_index is None:
        raise ValueError("Could not find SPACE section markers in INP file.")

    modified_inp = inp_data.copy()

    for i in range(start_index, end_index):
        line = modified_inp[i]

        if "EQUIPMENT-W/AREA" in line:
            indent = line[:line.index("EQUIPMENT-W/AREA")]
            
            # First, isolate EQUIPMENT-W/AREA and AREA/PERSON (if exists)
            equip_match = re.search(r'(EQUIPMENT-W/AREA\s*=\s*)\([^)]+\)', line)
            area_person_match = re.search(r'(AREA/PERSON\s*=\s*\{#PA\(".*?"\)\})', line)

            new_equip = f"{indent}{equip_match.group(1)}( {improvement:.2f} )\n" if equip_match else line

            if area_person_match:
                new_area = f"{indent}{area_person_match.group(1)}\n"
                modified_inp[i] = new_equip
                modified_inp.insert(i + 1, new_area)
            else:
                modified_inp[i] = new_equip

    return ''.join(modified_inp)
