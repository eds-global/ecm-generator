import os
import pandas as pd

def updateLPD(inp_data, light): 
    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "$ --"

    try:
        database = pd.read_excel("database/ML_ScaleUp_v02.xlsx", sheet_name='LightingPower-Improvement')
        improvement = database['Improvement'].iloc[light]
        # print(improvement)
    except FileNotFoundError:
        raise FileNotFoundError("Excel file not found: database/ML_ScaleUp_v02.xlsx")
    except ValueError:
        raise ValueError("Sheet 'LightingPower-Improvement' not found in the Excel file.")
    except IndexError:
        raise IndexError(f"Invalid 'light' index provided: {light}. Total improvements: {len(database)}")

    # if not os.path.isfile(inp_file_path):
    #     raise FileNotFoundError(f"Input file not found: {inp_file_path}")
    
    # with open(inp_file_path, 'r') as file:
    #     inp_data = file.readlines()

    start_index, end_index = None, None
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 4
            break

    for i, line in enumerate(inp_data):
        if end_marker in line:
            end_index = i
            if end_index > start_index:
                end_index = i - 4
                break
            else:
                continue
    
    if start_index is None or end_index is None:
        raise ValueError("Could not find SPACE section markers in INP file.")

    modified_inp = inp_data.copy()
    for i in range(start_index, end_index):
        if "LIGHTING-W/AREA" in modified_inp[i]:
            modified_inp[i] = f"   LIGHTING-W/AREA  = ( {improvement:.2f} )\n"

    return ''.join(modified_inp)