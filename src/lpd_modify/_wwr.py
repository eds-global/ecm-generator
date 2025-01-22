import streamlit as st
import pandas as pd
import numpy as np
import re
from src.lpd_modify import purge_windows

def extract_polygons(inp_file):
    with open(inp_file) as f:
        # Read all lines from the file and store them in a list named flist
        flist = f.readlines()
        
        # Initialize an empty list to store line numbers where 'Polygons' occurs
        polygon_count = [] 
        # Iterate through each line in flist along with its line number
        for num, line in enumerate(flist, 0):
            if 'Polygons' in line:
                polygon_count.append(num)
            if 'Wall Parameters' in line:
                numend = num
        # Store the line number of the first occurrence of 'Polygons'
        numstart = polygon_count[0] if polygon_count else None
        if not numstart:
            print("No 'Polygons' section found in the file.")
            return pd.DataFrame()  # Return an empty dataframe if no polygons section is found
        
        # Slice flist from the start of 'Polygons' to the line before 'Wall Parameters'
        polygon_rpt = flist[numstart:numend]
        
        # Initialize an empty dictionary to store polygon data
        polygon_data = {}
        current_polygon = None
        vertices = []
        
        # Iterate through the lines in polygon_rpt
        for line in polygon_rpt:
            if line.strip().startswith('"'):  # This indicates a new polygon
                if current_polygon:
                    polygon_data[current_polygon] = vertices
                current_polygon = line.split('"')[1].strip()  # Extract the polygon name
                vertices = []
            elif line.strip().startswith('V'):  # This is a vertex line
                try:
                    vertex = line.split('=')[1].strip()
                    vertex = tuple(map(float, vertex.strip('()').split(',')))
                    vertices.append(vertex)
                except ValueError:
                    pass  # Handle any lines that don't match the expected format
        if current_polygon:
            polygon_data[current_polygon] = vertices  # Add the last polygon

        # Debugging: Print the extracted polygon data
        # print("Extracted Polygon Data:")
        # print(polygon_data)
        
        # If polygon_data is empty, return an empty DataFrame
        if not polygon_data:
            print("No polygons data extracted.")
            return pd.DataFrame()
        
        # Get the maximum number of vertices in any polygon
        max_vertices = max(len(vertices) for vertices in polygon_data.values())

        # Create a DataFrame to store the polygon data
        result = []
        for polygon_name, vertices in polygon_data.items():
            # Fill missing vertex data with blanks
            vertices = list(vertices) + [''] * (max_vertices - len(vertices))
            result.append([polygon_name] + vertices)
        
        # Create the DataFrame and assign column names
        polygon_df = pd.DataFrame(result)
        column_names = ['Polygon'] + [f'V{i+1}' for i in range(max_vertices)]
        polygon_df.columns = column_names

        # Add a new column 'Total Vertices' to count non-empty vertices
        polygon_df['Total Vertices'] = polygon_df.iloc[:, 1:].apply(lambda row: sum(1 for v in row if v != ''), axis=1)

    return polygon_df

########################################################################

def extract_floor_space_wall_data(inp_file):
    # Initialize variables
    floor_data = []
    current_floor = None
    current_fh = None
    current_sh = None
    current_space = None
    current_polygon = None
    current_space_height = None
    walls_details = []
    inside_space_block = False  # Flag to track if inside a SPACE block

    # Helper function to append data
    def append_wall_data():
        for wall in walls_details:
            floor_data.append({
                'FLOOR': current_floor,
                'FLOOR-HEIGHT': current_fh,
                'SPACE-HEIGHT': current_sh,
                'SPACE': current_space,
                'SH2': current_space_height,
                'POLYGON': current_polygon,
                'EXTERIOR-WALL': wall.get('name'),
                'LOCATION': wall.get('location')
            })

    # Read input file
    with open(inp_file, 'r') as file:
        lines = file.readlines()

    for line in lines:
        line = line.strip()

        # Start FLOOR block
        if line.startswith('"') and '= FLOOR' in line:
            if walls_details:
                append_wall_data()
            # Reset for new floor
            current_floor = line.split('=')[0].strip().strip('"')
            current_fh = None
            current_sh = None
            current_space = None
            current_polygon = None
            current_space_height = None
            walls_details = []
            inside_space_block = False  # Reset SPACE block flag

        elif "FLOOR-HEIGHT" in line:
            current_fh = float(line.split('=')[1].strip())

        elif "SPACE-HEIGHT" in line:
            current_sh = float(line.split('=')[1].strip())

        # Start SPACE block
        elif line.startswith('"') and '= SPACE' in line:
            if walls_details:
                append_wall_data()
            # Reset for new space
            current_space = line.split('=')[0].strip().strip('"')
            current_polygon = None
            current_space_height = None
            walls_details = []
            inside_space_block = True  # Set SPACE block flag

        elif inside_space_block:
            # Capture SPACE block details
            if "HEIGHT" in line:
                current_space_height = float(line.split('=')[1].strip())
            elif "POLYGON" in line:
                current_polygon = line.split('=')[1].strip().strip('"')
            elif line == "..":  # End of SPACE block
                inside_space_block = False

        # Start EXTERIOR-WALL block
        elif line.startswith('"') and '= EXTERIOR-WALL' in line:
            wall_name = line.split('=')[0].strip().strip('"')
            walls_details.append({'name': wall_name, 'location': None})

        elif "LOCATION" in line and walls_details:
            walls_details[-1]['location'] = line.split('=')[1].strip()

        # End of a block
        elif line == "..":
            if walls_details:
                append_wall_data()
            # Reset after appending
            walls_details = []

    # Append remaining data
    if walls_details:
        append_wall_data()

    # Convert to DataFrame
    df = pd.DataFrame(floor_data)

    # Debug: Display extracted data
    print("Extracted data preview:")
    print(df.head())

    # Remove rows with LOCATION as 'TOP' or 'BOTTOM'
    df = df[~df['LOCATION'].isin(['TOP', 'BOTTOM'])]

    return df

def calculate_corr(row):
    try:
        # Attempt to split and calculate
        diff = row["Diff"].split(" - ")
        if len(diff) != 2:
            raise ValueError(f"Invalid Diff format: {row['Diff']}")
        col1, col2 = diff[0], diff[1]
        val1, val2 = row[col1], row[col2]
        return tuple(a - b for a, b in zip(val1, val2))
    except Exception as e:
        print(f"Error in row: {row.name}, Diff: {row['Diff']}, Error: {e}")
        return None  # Or a default value (e.g., (0, 0))

# Handle different data types in the Cordinate column
def calculate_distance(coord):
    if isinstance(coord, str) and coord.strip() != '':
        # If coord is a string and not blank, process it
        try:
            return np.sqrt(sum(float(num.strip())**2 for num in coord.strip('()').split(',')))
        except ValueError:
            return np.nan  # Handle invalid numeric values
    elif isinstance(coord, tuple):
        # If coord is a tuple, calculate the distance directly
        try:
            return np.sqrt(sum(float(num)**2 for num in coord))
        except ValueError:
            return np.nan
    else:
        # For other cases, return NaN
        return np.nan

def get_next_column(row):
    # Handle NaN or invalid cases
    if pd.isna(row["LOCATION"]) or not isinstance(row["LOCATION"], str):
        return np.nan
    
    if "V" in row["LOCATION"]:
        # Extract numeric part from 'SPACE-V1' or similar
        current_v = int(row["LOCATION"].split("-V")[1])
        total_vertices = row["Total Vertices"]
        
        if current_v < total_vertices:
            next_v = current_v + 1
            return f"SPACE-V{next_v}"
        elif current_v == total_vertices:
            return f"SPACE-V1"
    
    return ""

def create_ext_win(row):
    if isinstance(row, str):  # Process only strings
        # Replace 'Wall' with 'Win'
        new_value = row.replace('Wall', 'Win')
        # Append '.W1' after 'E<number>'
        new_value = re.sub(r'(E\d+)', r'\1.W1', new_value)
        return new_value
    return row  # Return the original value if it's not a string

def process_inp_file(inp_file):
    polygon_df = extract_polygons(inp_file)  # Polygon DataFrame
    df = extract_floor_space_wall_data(inp_file)  # Floor, Space, and Wall DataFrame
    df = pd.merge(df, polygon_df[['Polygon', 'Total Vertices']], left_on='POLYGON', right_on='Polygon', how='left')
    df.drop(columns=['Polygon'], inplace=True)
    return df

# df["Next-Column"] = df.apply(get_next_column, axis=1)
# df["Next-Column"] = df.apply(get_next_column, axis=1)
# df["EW-LOC_Num"] = df["LOCATION"].str.extract(r'(V\d+)')
# df["Next-Column_Num"] = df["Next-Column"].str.extract(r'(V\d+)')
# df["Diff"] = df["Next-Column_Num"] + " - " + df["EW-LOC_Num"]
# df["Diff"] = df["Diff"].fillna("")
# df = df.drop(columns=["EW-LOC_Num", "Next-Column_Num"])
# df['EXTERIOR-WINDOW'] = df['EXTERIOR-WALL'].apply(create_ext_win)
# columns = list(df.columns)
# ext_wall_index = columns.index('EXTERIOR-WALL')
# columns.insert(ext_wall_index + 1, columns.pop(columns.index('EXTERIOR-WINDOW')))
# df = df[columns]

# df = pd.merge(df, polygon_df, left_on='POLYGON', right_on='Polygon', how='inner')
# df = df.drop(columns=['Polygon'])

# df["Cordinate"] = df.apply(calculate_corr, axis=1)

# columns_to_remove = [f'V{i}' for i in range(1, 101)]
# df = df.drop(columns=[col for col in columns_to_remove if col in df.columns])

# df['D'] = df['Cordinate'].apply(lambda x: round(calculate_distance(x), 2) if x is not None else None)

# # Generate percentages from 10% to 90%
# df['SPACE-HEIGHT'] = pd.to_numeric(df['SPACE-HEIGHT'], errors='coerce')
# percentages = [i / 100 for i in range(5, 100, 5) if i == 5 or i % 10 == 0]
# percentages_SH = [i / 100 for i in range(5, 100, 5) if i == 5 or i % 10 == 0]

# for percent in percentages:
#     column_name = f"D{int(percent * 100)}%"  # Column names will be D10%, D20%, ..., D90%
#     df[column_name] = df['D'] * percent

# for percent in percentages_SH:
#     column_name = f"SPACE-HEIGHT{int(percent * 100)}%"  # Column names will be D10%, D20%, ..., D90%
#     df[column_name] = df['SPACE-HEIGHT'] * percent

# percentage_columns = [f"D{int(percent * 100)}%" for percent in percentages]
# percentage_columns_SH = [f"SPACE-HEIGHT{int(percent * 100)}%" for percent in percentages_SH]
# df[percentage_columns] = df[percentage_columns].round(2)
# df[percentage_columns_SH] = df[percentage_columns_SH].round(2)

# df['X'] = df['D5%']
# df['Y'] = df['SPACE-HEIGHT5%']

# for i, factor in enumerate([0.4, 0.5, 0.6, 0.7, 0.8, 0.9], start=1):
#     df[f'HEIGHT{i}'] = factor * df['SPACE-HEIGHT']
#     df[f'WIDTH{i}'] = factor * df['D']

# df['EXTERIOR-WINDOW'] = df['EXTERIOR-WALL'].apply(create_ext_win)
# columns = list(df.columns)
# ext_wall_index = columns.index('EXTERIOR-WALL')
# columns.insert(ext_wall_index + 1, columns.pop(columns.index('EXTERIOR-WINDOW')))
# df = df[columns]

# df.to_csv("window_coordinates.csv", index=False)
# purge_windows.process_all_inp_files_in_folder(inp_file_path, df)