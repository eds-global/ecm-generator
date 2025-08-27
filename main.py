import os, sys, time, webbrowser
import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image as PILImage
import numpy as np
import matplotlib.pyplot as plt
import json
import streamlit.components.v1 as components
import pdfplumber
import re
from io import BytesIO
import time
import random
from datetime import datetime
from shutil import copyfile
import glob as gb
import subprocess
import shutil
from pathlib import Path
from collections import defaultdict
import streamlit.web.cli as stcli
from src import insertWall, insertConst, orient, lighting, equip, windows, insertRoof, wwr
import traceback
from helper import *
from report_ext import *
from src import lv_b, ls_c, lv_d, pv_a_loop, sv_a, beps, bepu, lvd_summary, sva_zone, locationInfo, masterFile, sva_sys_type, pv_a_pump, pv_a_heater, pv_a_equip, pv_a_tower, ps_e, inp_shgc


# --- Streamlit Page Config ---
st.set_page_config(page_title="ECM Batch Run",page_icon="💡", layout='wide',)

# --- Inject Custom CSS ---
st.markdown("""
    <style>
        .block-container { padding-top: 0rem !important; }
        header, main { margin-top: 0rem !important; padding-top: 0rem !important; }
        .stButton>button { box-shadow: 1px 1px 1px rgba(0, 0, 0, 0.8); }
        .heading-with-shadow {
            text-align: left;
            color: red;
            text-shadow: 0px 8px 4px rgba(255, 255, 255, 0.4);
            background-color: white;
        }
        body {
            background-color: #bfe1ff;
            animation: changeColor 5s infinite;
        }
        #MainMenu, footer, .viewerBadge_container__1QSob {visibility: hidden;}
        header .stApp [title="View source on GitHub"] { display: none; }
        .stApp header, .stApp footer {visibility: hidden;}
        .stButton button {
            height: 30px;
            width: 166px;
        }
    </style>
""", unsafe_allow_html=True)


# function to get report in csv and save in specific folders
def get_report_and_save(report_function, name, file_suffix, folder_name, path):
    try:
        print("calling report function: ", report_function, name, path)
        report = report_function(name, path)
    except:
        print("Skipping...")
    # get file path name as .csv
    file_path = os.path.join(path, f'{x[0]}_{file_suffix}.csv')
    # if that file already exist, replace with other file.
    if os.path.isfile(file_path):
        os.remove(file_path)
    # writing csv file with headers and no index column
    print("Done-", name)
    with open(file_path, 'w', newline='') as f:
        report.to_csv(f, header=True, index=False, mode='wt') 

def remove_section_content(content, start_marker, end_marker):
    """
    Function to remove content between start_marker and end_marker.
    Deletes lines between start_marker_index + 3 and end_marker_index - 3.
    """
    new_content = []
    
    start_index = -1
    end_index = -1
    for i, line in enumerate(content):
        if start_marker in line and start_index == -1:
            start_index = i
        if end_marker in line and end_index == -1:
            end_index = i
    
    if start_index != -1 and end_index != -1:
        new_content.extend(content[:start_index + 1])
        new_content.extend(content[start_index + 1:start_index + 3])
        new_content.extend(content[end_index - 3:end_index])
        new_content.append(content[end_index])
        new_content.extend(content[end_index + 1:])
    else:
        new_content = content
    
    return new_content

def delete_glass_type_codes(content):
    return remove_section_content(content, "$              Glass Type Codes", "$              Glass Types")

def modify_glass_types(content, start_marker, end_marker):
    """
    Modify GLASS-TYPE names by adding 'ML_' as a prefix.
    """
    start_index = None
    end_index = None

    # Identify the start and end of the Glass Types section
    for i, line in enumerate(content):
        if start_marker in line and start_index is None:
            start_index = i
        if end_marker in line and start_index is not None:
            end_index = i
            break

    if start_index is None or end_index is None:
        # Markers not found, return content unchanged
        return content

    # Modify GLASS-TYPE names in the Glass Types section
    for i in range(start_index + 1, end_index):
        if "= GLASS-TYPE" in content[i]:
            parts = content[i].split("=")
            if len(parts) > 1:
                glass_type_name = parts[0].strip().strip('"')
                modified_name = f'ML_{glass_type_name}'
                content[i] = content[i].replace(glass_type_name, modified_name)

    return content, glass_type_name

def delete_window_layers(content):
    return remove_section_content(content, "$              Window Layers", "$              Lamps / Luminaries / Lighting Systems")

def remove_window_sections(content, start_marker, end_marker):
    """
    Removes all = WINDOW sections between the start_marker and end_marker.
    """
    start_index = None
    end_index = None
    for i, line in enumerate(content):
        if start_marker in line and start_index is None:
            start_index = i
        if end_marker in line and start_index is not None:
            end_index = i
            break
    
    if start_index is None or end_index is None:
        # Markers not found, return content unchanged
        return content

    # Extract the content between the markers
    pre_marker_content = content[:start_index + 1]
    between_marker_content = content[start_index + 1:end_index]
    post_marker_content = content[end_index:]
    
    # Filter out = WINDOW sections
    filtered_content = []
    skip_window_section = False

    for line in between_marker_content:
        if "= WINDOW" in line:
            skip_window_section = True
        if skip_window_section and line.strip() == "..":
            skip_window_section = False
            continue
        if not skip_window_section:
            filtered_content.append(line)

    return pre_marker_content + filtered_content + post_marker_content

def include_window_sections(content, start_marker, end_marker, df, glass_type_name, height):
    # Initialize indices to None
    start_index = None
    end_index = None

    # Search for the markers in the content (which is a list of lines)
    for i, line in enumerate(content):
        if start_marker in line and start_index is None:
            start_index = i 
        if end_marker in line and start_index is not None:
            end_index = i
            break

    if start_index is None or end_index is None:
        return content  # If markers are not found, return the original content

    # Extract the content between the markers
    section = content[start_index:end_index]

    # Prepare the modified section
    modified_section = []

    # Iterate through the section to process EXTERIOR-WALL entries
    current_wall_name = None
    include_window = False

    for line in section:
        modified_section.append(line)

        if "= EXTERIOR-WALL" in line:
            current_wall_name = line.split("=")[0].strip().strip('"')  # Extract wall name
            include_window = True  # Allow processing unless LOCATION says otherwise

        if "LOCATION" in line and current_wall_name:
            location_value = line.split("=")[1].strip()
            if location_value in ["TOP", "BOTTOM"]:
                include_window = False  # Skip if LOCATION is TOP or BOTTOM
                current_wall_name = None  # Reset wall name for safety

        if include_window and line.strip() == "..":
            if current_wall_name in df['EXTERIOR-WALL'].values:
                row = df[df['EXTERIOR-WALL'] == current_wall_name].iloc[0]
                # Construct window name
                base_name = current_wall_name.split("Wall")[0].strip()  # Ensure "Wall" is excluded
                identifier = current_wall_name.split("(")[1].strip(")")  # Extract the identifier
                window_name = f"{base_name} Win ({identifier}.W1)"  # Construct the correct name
                window_section = f'''"{window_name}" = WINDOW
   GLASS-TYPE       = "ML_{glass_type_name}"
   FRAME-WIDTH      = 0
   X                = {row['X']}
   Y                = {row['Y']}
   HEIGHT           = {row[f'HEIGHT{height + 1}']}
   WIDTH            = {row[f'WIDTH{height + 1}']}
   FRAME-CONDUCT    = 2.781
   ..
'''
                modified_section.append(window_section)  # Insert window section
                include_window = False  # Reset flag

    # Reassemble content with the modified section
    modified_content = content[:start_index] + modified_section + content[end_index:]
    return modified_content

def process_sections(content, df, height):
    # df.to_csv("window_coordinates.csv", index=False)
    df = df[df['SH2'].isna() | (df['SH2'] == '')]
    # with open(file_path, 'r') as file:
    #     content = file.readlines()
    
    content = delete_glass_type_codes(content)
    content, glass_type_name = modify_glass_types(content, "$              Glass Types", "$              Window Layers")
    content = delete_window_layers(content)
    content = remove_window_sections(content, "$ **      Floors / Spaces / Walls / Windows / Doors      **",
        "$ **              Electric & Fuel Meters                 **")
    content = include_window_sections(content, "$ **      Floors / Spaces / Walls / Windows / Doors      **",
    "$ **              Electric & Fuel Meters                 **", df, glass_type_name, height)  # or any height value you want

    
    # dir_name, file_name = os.path.split(file_path)
    # modified_file_name = 'Purged_90%_' + file_name
    # modified_file_path = os.path.join(dir_name, modified_file_name)

    return content

def process_all_inp_files_in_folder(inp_path, df, height):
    # print(inp_path)
    """Process all .inp files in a folder, modifying each one using the process_sections function."""
    process_sections(inp_path, df, height)

def extract_polygons(flist):
    # with open(inp_file) as f:
    #     # Read all lines from the file and store them in a list named flist
    #     flist = f.readlines()
        
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
        # print("No 'Polygons' section found in the file.")
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
        # print("No polygons data extracted.")
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

def extract_floor_space_wall_data(lines):
    import pandas as pd

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
    # with open(inp_file, 'r') as file:
    #     lines = file.readlines()

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
    # print("Extracted data preview:")
    # print(df.head())

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
        # print(f"Error in row: {row.name}, Diff: {row['Diff']}, Error: {e}")
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

def resource_path(relative_path):
    """Get absolute path to resource (works for dev and PyInstaller exe)"""
    if getattr(sys, 'frozen', False):  # Running in exe
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def main():
    # Heading
    _, col2, _ = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <h2 style='text-align: center; white-space: nowrap;'>
            <span style='color:#FF5733;'>𝓡𝓾𝓷 𝓢𝓲𝓶𝓾𝓵𝓪𝓽𝓲𝓸𝓷 𝓸𝓯 𝓮𝓠𝓤𝓔𝓢𝓣</span>
        </h2>
        """, unsafe_allow_html=True)

    st.markdown('<hr style="border:1px solid black">', unsafe_allow_html=True)

    # Load location database
    csv_path = resource_path(os.path.join("database", "Simulation_locations.csv"))
    weather_df = pd.read_csv(csv_path)
    db_path = resource_path(os.path.join("database", "AllData.xlsx"))
    output_csv = resource_path(os.path.join("database", "Randomized_Sheet.xlsx"))
    updated_df = pd.read_excel(output_csv)

    # Inputs
    user_nm = "Rajeev"
    col1, col2, col3 = st.columns(3)
    locations = ["NewDelhi", "Ahemdabad", "Bangalore", "Chennai"]

    with col1:
        project_name = st.text_input("📝 Project Name", placeholder="Enter project name")
        st.markdown("""
            <style>
            div[data-testid="stFileUploader"] section {
                padding: 0.25rem 0 !important;  /* Reduce vertical padding */
            }
            div[data-testid="stFileUploader"] div[role="button"] {
                padding: 0.2rem 0.5rem !important;  /* Reduce height of clickable area */
                font-size: 0.85rem !important;      /* Smaller text */
            }
            </style>
        """, unsafe_allow_html=True)
        uploaded_file = st.file_uploader("📤 Upload a single eQUEST INP file", type=["inp"])

    with col2:
        user_input = st.selectbox("🌍 Select Location", locations).strip().lower()
    
    with col3:
        typology = st.text_input("🏠 Enter Building Typology", placeholder="Enter Typology")
    
    # with col4:
    if uploaded_file:
        # Save uploaded file in current working directory (same as your app folder)
        uploaded_inp_path = os.path.join(os.getcwd(), uploaded_file.name)
        with open(uploaded_inp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Output folder = same folder where uploaded file is saved
        output_inp_folder = os.path.dirname(uploaded_inp_path)
        inp_folder = output_inp_folder
            # st.text_input("📁 Output Folder Path", value=output_inp_folder, disabled=True)

    # with col4:
    #     output_inp_folder = st.text_input("📁 Output Folder Path", placeholder="Type or paste the output folder path")
    #     inp_folder = output_inp_folder
    #     if uploaded_file:
    #         os.makedirs(inp_folder, exist_ok=True)
    #         uploaded_inp_path = os.path.join(inp_folder, uploaded_file.name)
            
    run_cnt = 1
    location_id, weather_path = "", ""
    matched_row = weather_df[weather_df['Sim_location'].str.lower().str.contains(user_input)]
    if not matched_row.empty:
        location_id = matched_row.iloc[0]['Location_ID']
        weather_path = matched_row.iloc[0]['Weather_file']
    elif user_input:
        st.error("❌ Location not found.")

    # Start simulation
    if st.button("🚀 Start Simulation"):
        os.makedirs(output_inp_folder, exist_ok=True)
        new_batch_id = f"{int(time.time())}"  # unique ID

        selected_rows = updated_df[updated_df['Batch_ID'] == run_cnt]
        batch_output_folder = os.path.join(output_inp_folder, f"{user_nm}_Batch_{new_batch_id}")
        os.makedirs(batch_output_folder, exist_ok=True)

        num = 1
        modified_files = []
        for _, row in selected_rows.iterrows():
            selected_inp = row["Selected_INP"]
            new_inp_name = f"{row['Wall']}_{row['Roof']}_{row['Glazing']}_{row['Orient']}_{row['Light']}_{row['WWR']}_{row['Equip']}_{selected_inp}"
            new_inp_path = os.path.join(batch_output_folder, new_inp_name)

            inp_file_path = os.path.join(inp_folder, selected_inp)
            if not os.path.exists(inp_file_path):
                st.error(f"File {inp_file_path} not found. Skipping.")
                continue

            # st.info(f"Modifying INP file {num}: {selected_inp} -> {new_inp_name}")
            num += 1

            # Apply modifications
            inp_content = wwr.process_window_insertion_workflow(inp_file_path, row["WWR"] + 1)
            inp_content = orient.updateOrientation(inp_content, row["Orient"])
            inp_content = lighting.updateLPD(inp_content, row['Light'])
            inp_content = insertWall.update_Material_Layers_Construction(inp_content, row["Wall"])
            inp_content = insertRoof.update_Material_Layers_Construction(inp_content, row["Roof"])
            inp_content = insertRoof.removeDuplicates(inp_content)
            inp_content = equip.updateEquipment(inp_content, row['Equip'])
            inp_content = windows.insert_glass_types_multiple_outputs(inp_content, row['Glazing'])

            with open(new_inp_path, 'w') as file:
                file.writelines(inp_content)
            modified_files.append(new_inp_name)
        st.markdown(
            "<strong>🔧 Modified Files →</strong> " +
            "".join([
                f"<span style='background:#e0f7fa; color:#00796b; padding:4px 8px; "
                f"border-radius:12px; margin:2px; display:inline-block;'>"
                f"{name}</span>"
                for name in modified_files
            ]),
            unsafe_allow_html=True
        )


        simulate_files = []
        # Copy and run batch script
        if uploaded_file is None:
            st.error("Please upload an INP file before starting the simulation.")
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            shutil.copy(os.path.join(script_dir, "script.bat"), batch_output_folder)
            inp_files = [f for f in os.listdir(batch_output_folder) if f.lower().endswith(".inp")]
            for inp_file in inp_files:
                file_path = os.path.join(batch_output_folder, os.path.splitext(inp_file)[0])
                subprocess.call(
                    [os.path.join(batch_output_folder, "script.bat"), file_path, weather_path],
                    shell=True
                )
                simulate_files.append(inp_file)
            st.markdown(
                "<strong>⚡ Simulated Files →</strong> " +
                "".join([
                    f"<span style='background:#e0f7fa; color:#00796b; padding:4px 8px; "
                    f"border-radius:12px; margin:2px; display:inline-block;'>"
                    f"{name}</span>"
                    for name in simulate_files
                ]),
                unsafe_allow_html=True
            )
            # subprocess.call([os.path.join(batch_output_folder, "script.bat"), batch_output_folder, weather_path], shell=True)
            
            required_sections = ['BEPS', 'BEPU', 'LS-C', 'LV-B', 'LV-D', 'PS-E', 'SV-A']
            log_file_path = check_missing_sections(batch_output_folder, required_sections, new_batch_id, user_nm)
            get_failed_simulation_data(batch_output_folder, log_file_path)
            clean_folder(batch_output_folder)
            get_files_for_data_extraction(batch_output_folder, log_file_path, new_batch_id, location_id, user_nm)
        logFile = pd.read_excel(log_file_path)
        total_runs = len(updated_df)
        total_sims = len(logFile)
        success_count = (logFile["Status"] == "Success").sum()
        success_rate = (success_count / total_sims) * 100 if total_sims > 0 else 0
        # 5 equal columns
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.markdown(
                "<div style='background:#e8f5e9; padding:15px; border-radius:12px; text-align:center;'>"
                "✅ <br><b>Completed</b><br>Simulation & Extraction</div>",
                unsafe_allow_html=True
            )

        with col2:
            st.markdown(
                "<div style='background:#f3e5f5; padding:15px; border-radius:12px; text-align:center;'>"
                "📊 <b>Log File</b></div>",
                unsafe_allow_html=True
            )
            with st.expander("🔽 Click to View Log File"):
                st.dataframe(logFile, use_container_width=True)
                
        with col3:
            st.markdown(
                f"<div style='background:#e3f2fd; padding:15px; border-radius:12px; text-align:center;'>"
                f"🧮 <br><b>Total Runs</b><br>{total_runs}</div>",
                unsafe_allow_html=True
            )

        with col4:
            st.markdown(
                f"<div style='background:#ede7f6; padding:15px; border-radius:12px; text-align:center;'>"
                f"🧮 <br><b>Total Sims</b><br>{total_sims}</div>",
                unsafe_allow_html=True
            )

        with col5:
            st.markdown(
                f"<div style='background:#fff3e0; padding:15px; border-radius:12px; text-align:center;'>"
                f"📈 <br><b>Success Rate</b><br>{success_rate:.2f}%</div>",
                unsafe_allow_html=True
            )

if __name__ == "__main__":
    main()
