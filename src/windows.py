
import pandas as pd
import re

######################################################
##################  Glazing R-Value ##################
######################################################

def insert_glass_UVal(inp_data, row_num):
    if row_num == 0:
        return inp_data
    
    start_marker = "Glass Types"
    end_marker = "Window Layers"

    df = pd.read_excel("database/AllData.xlsx", sheet_name="GlazedR")

    if "U-Value" not in df.columns:
        raise ValueError("U-Value column not found")

    if row_num >= len(df):
        raise IndexError("U_val out of range")

    u_value = df.loc[row_num, "U-Value"]

    start_idx = end_idx = None
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_idx = i
        elif end_marker in line and start_idx is not None:
            end_idx = i
            break

    if start_idx is None or end_idx is None:
        print("Glass Types section not found")
        return inp_data

    for i in range(start_idx, end_idx):
        if "GLASS-CONDUCT" in inp_data[i] and "=" in inp_data[i]:
            inp_data[i] = re.sub(
                r"(GLASS-CONDUCT\s*=\s*)([0-9.]+)",
                rf"\g<1>{u_value}",
                inp_data[i]
            )

    return inp_data
    

def readSCUVal(inp_data, row_num=0):
    if row_num > 0:
        return inp_data
    # Convert string to list if needed
    if isinstance(inp_data, str):
        inp_data = inp_data.splitlines()

    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    start_index, end_index = None, None
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 1
        if end_marker in line:
            end_index = i - 1
            break

    if start_index is None or end_index is None:
        print("WINDOW section not found.")
        return pd.DataFrame()

    section_lines = inp_data[start_index:end_index]

    windows = []
    current_window = None

    for line in section_lines:

        # Start of WINDOW object
        if "= WINDOW" in line:
            if current_window:
                windows.append(current_window)

            name = re.search(r'"(.*?)"\s*=\s*WINDOW', line)
            current_window = {
                "Name": name.group(1) if name else None,
                "WindowType": None,
                "FrameWidth": None,
                "X": None,
                "Y": None,
                "Height": None,
                "Width": None,
                "FrameConduct": None
            }

        elif current_window:
            # Extract parameters
            if "WINDOW-TYPE" in line:
                current_window["WindowType"] = extract_value(line)

            elif "FRAME-WIDTH" in line:
                current_window["FrameWidth"] = extract_value(line)

            elif re.match(r"\s*X\s*=", line):
                current_window["X"] = extract_value(line)

            elif re.match(r"\s*Y\s*=", line):
                current_window["Y"] = extract_value(line)

            elif "HEIGHT" in line:
                current_window["Height"] = extract_value(line)

            elif "WIDTH" in line and "FRAME-WIDTH" not in line:
                current_window["Width"] = extract_value(line)

            elif "FRAME-CONDUCT" in line:
                current_window["FrameConduct"] = extract_value(line)

    # Append last window
    if current_window:
        windows.append(current_window)
    df_windows = pd.DataFrame(windows)
    # df_windows.to_csv("output.csv", index=False)

def extract_value(line):
    """Extract numeric or string value after '='"""
    value = line.split("=", 1)[1].strip().strip(",")
    value = value.replace('"', '')
    try:
        return float(value)
    except ValueError:
        return value

def insert_glass_types_multiple_outputs(inp_data, row_glaze):
    # If inp_data is a string, convert to list of lines
    if isinstance(inp_data, str):
        inp_data = inp_data.splitlines(keepends=True)

    # if row_wall == 0 and row_roof == 0 and row_equip == 0 and row_light == 0 and row_wwr == 0 and row_glaze == 0:
    #     if isinstance(inp_data, str):
    #         inp_data = inp_data.splitlines()

    #     start_marker = "Floors / Spaces / Walls / Windows / Doors"
    #     end_marker = "Electric & Fuel Meters"

    #     start_index, end_index = None, None
    #     for i, line in enumerate(inp_data):
    #         if start_marker in line:
    #             start_index = i + 1
    #         if end_marker in line:
    #             end_index = i - 1
    #             break

    #     if start_index is None or end_index is None:
    #         print("WINDOW section not found.")
    #         return pd.DataFrame()

    #     section_lines = inp_data[start_index:end_index]

    #     windows = []
    #     current_window = None

    #     for line in section_lines:

    #         # Start of WINDOW object
    #         if "= WINDOW" in line:
    #             if current_window:
    #                 windows.append(current_window)

    #             name = re.search(r'"(.*?)"\s*=\s*WINDOW', line)
    #             current_window = {
    #                 "Name": name.group(1) if name else None,
    #                 "WindowType": None,
    #                 "FrameWidth": None,
    #                 "X": None,
    #                 "Y": None,
    #                 "Height": None,
    #                 "Width": None,
    #                 "FrameConduct": None
    #             }

    #         elif current_window:
    #             # Extract parameters
    #             if "GLASS-TYPE" in line or "WINDOW-TYPE" in line:
    #                 current_window["WindowType"] = extract_value(line)

    #             elif "FRAME-WIDTH" in line:
    #                 current_window["FrameWidth"] = extract_value(line)

    #             elif re.match(r"\s*X\s*=", line):
    #                 current_window["X"] = extract_value(line)

    #             elif re.match(r"\s*Y\s*=", line):
    #                 current_window["Y"] = extract_value(line)

    #             elif "HEIGHT" in line:
    #                 current_window["Height"] = extract_value(line)

    #             elif "WIDTH" in line and "FRAME-WIDTH" not in line:
    #                 current_window["Width"] = extract_value(line)

    #             elif "FRAME-CONDUCT" in line:
    #                 current_window["FrameConduct"] = extract_value(line)

    #     # Append last window
    #     if current_window:
    #         windows.append(current_window)
    #     df_windows = pd.DataFrame(windows)

    if row_glaze == 0:
        return inp_data
        
    start_marker = "Glass Types"
    end_marker = "Window Layers"

    df = pd.read_excel("database/AllData.xlsx", sheet_name="Glazing")

    if "SC" not in df.columns:
        raise ValueError("SC column not found")

    if row_glaze >= len(df):
        raise IndexError("glaze_val out of range")

    sc_value = df.loc[row_glaze, "SC"]

    start_idx = end_idx = None
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_idx = i
        elif end_marker in line and start_idx is not None:
            end_idx = i
            break

    if start_idx is None or end_idx is None:
        print("Glass Types section not found")
        return inp_data

    for i in range(start_idx, end_idx):
        if "SHADING-COEF" in inp_data[i] and "=" in inp_data[i]:
            inp_data[i] = re.sub(
                r"(SHADING-COEF\s*=\s*)([0-9.]+)",
                rf"\g<1>{sc_value}",
                inp_data[i]
            )

    return inp_data




# import pandas as pd
# import re

# def readSCUVal(inp_data, row_num=0):
#     if row_num > 0:
#         return inp_data
#     # Convert string to list if needed
#     if isinstance(inp_data, str):
#         inp_data = inp_data.splitlines()

#     start_marker = "Floors / Spaces / Walls / Windows / Doors"
#     end_marker = "Electric & Fuel Meters"

#     start_index, end_index = None, None
#     for i, line in enumerate(inp_data):
#         if start_marker in line:
#             start_index = i + 1
#         if end_marker in line:
#             end_index = i - 1
#             break

#     if start_index is None or end_index is None:
#         print("WINDOW section not found.")
#         return pd.DataFrame()

#     section_lines = inp_data[start_index:end_index]

#     windows = []
#     current_window = None

#     for line in section_lines:

#         # Start of WINDOW object
#         if "= WINDOW" in line:
#             if current_window:
#                 windows.append(current_window)

#             name = re.search(r'"(.*?)"\s*=\s*WINDOW', line)
#             current_window = {
#                 "Name": name.group(1) if name else None,
#                 "WindowType": None,
#                 "FrameWidth": None,
#                 "X": None,
#                 "Y": None,
#                 "Height": None,
#                 "Width": None,
#                 "FrameConduct": None
#             }

#         elif current_window:
#             # Extract parameters
#             if "WINDOW-TYPE" in line:
#                 current_window["WindowType"] = extract_value(line)

#             elif "FRAME-WIDTH" in line:
#                 current_window["FrameWidth"] = extract_value(line)

#             elif re.match(r"\s*X\s*=", line):
#                 current_window["X"] = extract_value(line)

#             elif re.match(r"\s*Y\s*=", line):
#                 current_window["Y"] = extract_value(line)

#             elif "HEIGHT" in line:
#                 current_window["Height"] = extract_value(line)

#             elif "WIDTH" in line and "FRAME-WIDTH" not in line:
#                 current_window["Width"] = extract_value(line)

#             elif "FRAME-CONDUCT" in line:
#                 current_window["FrameConduct"] = extract_value(line)

#     # Append last window
#     if current_window:
#         windows.append(current_window)
#     df_windows = pd.DataFrame(windows)
#     df_windows.to_csv("output.csv", index=False)

# def extract_value(line):
#     """Extract numeric or string value after '='"""
#     value = line.split("=", 1)[1].strip().strip(",")
#     value = value.replace('"', '')
#     try:
#         return float(value)
#     except ValueError:
#         return value

# def insert_glass_types_multiple_outputs(inp_data, row_wall, row_roof, row_light, row_equip, row_glaze, row_wwr):
#     # If inp_data is a string, convert to list of lines
#     if isinstance(inp_data, str):
#         inp_data = inp_data.splitlines(keepends=True)

#     if row_wall == 0 and row_roof == 0 and row_equip == 0 and row_light == 0 and row_wwr == 0 and row_glaze == 0:
#         if isinstance(inp_data, str):
#             inp_data = inp_data.splitlines()

#         start_marker = "Floors / Spaces / Walls / Windows / Doors"
#         end_marker = "Electric & Fuel Meters"

#         start_index, end_index = None, None
#         for i, line in enumerate(inp_data):
#             if start_marker in line:
#                 start_index = i + 1
#             if end_marker in line:
#                 end_index = i - 1
#                 break

#         if start_index is None or end_index is None:
#             print("WINDOW section not found.")
#             return pd.DataFrame()

#         section_lines = inp_data[start_index:end_index]

#         windows = []
#         current_window = None

#         for line in section_lines:

#             # Start of WINDOW object
#             if "= WINDOW" in line:
#                 if current_window:
#                     windows.append(current_window)

#                 name = re.search(r'"(.*?)"\s*=\s*WINDOW', line)
#                 current_window = {
#                     "Name": name.group(1) if name else None,
#                     "WindowType": None,
#                     "FrameWidth": None,
#                     "X": None,
#                     "Y": None,
#                     "Height": None,
#                     "Width": None,
#                     "FrameConduct": None
#                 }

#             elif current_window:
#                 # Extract parameters
#                 if "GLASS-TYPE" in line or "WINDOW-TYPE" in line:
#                     current_window["WindowType"] = extract_value(line)

#                 elif "FRAME-WIDTH" in line:
#                     current_window["FrameWidth"] = extract_value(line)

#                 elif re.match(r"\s*X\s*=", line):
#                     current_window["X"] = extract_value(line)

#                 elif re.match(r"\s*Y\s*=", line):
#                     current_window["Y"] = extract_value(line)

#                 elif "HEIGHT" in line:
#                     current_window["Height"] = extract_value(line)

#                 elif "WIDTH" in line and "FRAME-WIDTH" not in line:
#                     current_window["Width"] = extract_value(line)

#                 elif "FRAME-CONDUCT" in line:
#                     current_window["FrameConduct"] = extract_value(line)

#         # Append last window
#         if current_window:
#             windows.append(current_window)
#         df_windows = pd.DataFrame(windows)
#         df_windows.to_csv("output.csv", index=False)
        
#     # Define section markers
#     start_marker = "Glass Types"
#     end_marker = "Window Layers"
#     start_marker1 = "Floors / Spaces / Walls / Windows / Doors"
#     end_marker1 = "Electric & Fuel Meters"

#     # Extract the climate from the file name (simulate path for climate extraction)
#     inp_file_name = "example_Pune.inp"  # replace with actual filename logic if needed
#     climate = inp_file_name.split('_')[-1].replace('.inp', '')  # e.g., 'Pune'

#     # Load and filter the database
#     database_path = 'database/AllData.xlsx'
#     database = pd.read_excel(database_path, sheet_name='Glazing')
#     filtered_db = database[database['Climate'] == climate]
#     if filtered_db.empty:
#         filtered_db = database[database['Climate'] == 'Others']

#     # Validate row number
#     if row_glaze >= len(filtered_db):
#         print(f"Row number {row_glaze} out of range for climate '{climate}' in database.")
#         return inp_data

#     # Find Glass Types insertion section
#     start_index, end_index = None, None
#     for i, line in enumerate(inp_data):
#         if start_marker in line:
#             start_index = i + 1
#         if end_marker in line:
#             end_index = i
#             break
#     if start_index is None or end_index is None:
#         print("Glass Types section not found.")
#         return inp_data

#     # Find Windows section
#     start_index1, end_index1 = None, None
#     for i, line in enumerate(inp_data):
#         if start_marker1 in line:
#             start_index1 = i + 4
#         if end_marker1 in line:
#             end_index1 = i - 4
#             break
#     if start_index1 is None or end_index1 is None:
#         print("Windows section not found.")
#         return inp_data

#     # Clean unwanted lines in WINDOW section
#     updated_data1 = []
#     in_window_section = False
#     for i in range(start_index1, end_index1 + 1):
#         line = inp_data[i].strip()
#         if line.endswith("= WINDOW"):
#             in_window_section = True
#             updated_data1.append(inp_data[i])
#             continue
#         if in_window_section:
#             if any(key in line for key in ["FRAME-WIDTH", "FRAME-CONDUCT", "SPACER-TYPE"]):
#                 continue
#             if line == "..":
#                 in_window_section = False
#         updated_data1.append(inp_data[i])
#     inp_data = inp_data[:start_index1] + updated_data1 + inp_data[end_index1 + 1:]

#     # Get selected row
#     selected_row = filtered_db.iloc[row_glaze]

#     # Prepare insertion content
#     insertion_lines = [
#         '$ ---------------------------------------------------------\n\n',
#         f'"{selected_row["GLASS-TYPE"]}" = GLASS-TYPE\n',
#         '   TYPE             = SHADING-COEF\n',
#         f'   SHADING-COEF     = {round(selected_row["SHADING-COEF"],2)}\n',
#         f'   GLASS-CONDUCT    = {round(selected_row["GLASS-CONDUCT"],2)}\n',
#         f'   VIS-TRANS        = {selected_row["VIS-TRANS"]}\n',
#         '   ..\n\n',
#         '$ ---------------------------------------------------------\n'
#     ]

#     # Insert glass type lines
#     updated_data = inp_data[:start_index] + insertion_lines + inp_data[end_index:]

#     # Update GLASS-TYPE in WINDOW sections
#     in_window_section = False
#     for i in range(start_index1, end_index1 + 1):
#         line = updated_data[i].strip()
#         if line.endswith("= WINDOW"):
#             in_window_section = True
#         elif in_window_section and line.startswith("GLASS-TYPE"):
#             updated_data[i] = f'   GLASS-TYPE       = "{selected_row["GLASS-TYPE"]}"\n'
#             in_window_section = False

#     return updated_data
