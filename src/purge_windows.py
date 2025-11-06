import os
import pandas as pd
import re

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

def include_window_sections(content, start_marker, end_marker, df, glass_type_name):
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
   HEIGHT           = {row['HEIGHT6']}
   WIDTH            = {row['WIDTH6']}
   FRAME-CONDUCT    = 2.781
   ..
'''
                modified_section.append(window_section)  # Insert window section
                include_window = False  # Reset flag

    # Reassemble content with the modified section
    modified_content = content[:start_index] + modified_section + content[end_index:]
    return modified_content

def process_sections(file_path, df):
    # df.to_csv("window_coordinates.csv", index=False)
    df = df[df['SH2'].isna() | (df['SH2'] == '')]
    with open(file_path, 'r') as file:
        content = file.readlines()
    
    content = delete_glass_type_codes(content)
    content, glass_type_name = modify_glass_types(content, "$              Glass Types", "$              Window Layers")
    content = delete_window_layers(content)
    content = remove_window_sections(content, "$ **      Floors / Spaces / Walls / Windows / Doors      **",
        "$ **              Electric & Fuel Meters                 **")
    content = include_window_sections(content, "$ **      Floors / Spaces / Walls / Windows / Doors      **",
        "$ **              Electric & Fuel Meters                 **", df, glass_type_name)
    
    dir_name, file_name = os.path.split(file_path)
    modified_file_name = 'Purged_90%_' + file_name
    modified_file_path = os.path.join(dir_name, modified_file_name)

    with open(modified_file_path, 'w') as file:
        file.writelines(content)
    df.to_csv("Testing.csv", index=False)
    print(f"Modified file saved as: {modified_file_path}")

def process_all_inp_files_in_folder(inp_path, df):
    print(inp_path)
    """
    Process all .inp files in a folder, modifying each one using the process_sections function.
    """
    if inp_path.endswith('.inp'):
        process_sections(inp_path, df)