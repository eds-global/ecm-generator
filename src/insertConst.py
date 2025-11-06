import pandas as pd
import os
import re

def extract_value(s):
    # Define the pattern to match
    pattern = r'(?:\\[^\\]*){6}\\([^\\-]*)-'
    # Search for the pattern in the string
    match = re.search(pattern, s)
    # If match found, return the captured group
    if match:
        return match.group(1)
    else:
        return None

def update_external_wall(name):
    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    with open(name, 'r') as file:
        data = file.readlines()

    start_index = None
    end_index = None

    # Finding start and end indices
    for i, line in enumerate(data):
        if start_marker in line:
            start_index = i
        if end_marker in line:
            end_index = i
            break

    if start_index is not None and end_index is not None:
        for i in range(start_index + 4, end_index - 3):
            if "EXTERIOR-WALL" in data[i]:
                if "LOCATION         = TOP" in data[i + 2]:
                    pass  # Do nothing, leave it unchanged
                elif "LOCATION         = BOTTOM" in data[i + 2]:
                    pass  # Do nothing, leave it unchanged
                else:
                    # Extract value from the name
                    value = extract_value(name)
                    if value is not None:  # Check if value is not None
                        # Replace TOP or BOTTOM with the extracted value
                        location_index = data[i + 1].index("=") + 1
                        if "TOP" not in data[i + 2]:
                            data[i + 1] = data[i + 1][:location_index] + ' "' + value + '"\n'
                        elif "BOTTOM" not in data[i + 2]:
                            data[i + 1] = data[i + 1][:location_index] + ' "' + value + '"\n'

        
    # Write modified data back to the file
    with open(name, 'w') as file:
        file.writelines(data)
