import streamlit as st

def getOrientation(data, orient):
    # Open and read the input file
    # with open(data_path, 'r') as file:
    #     data = file.readlines()
    
    if orient is None:
        return data
        
    start_marker = "Site and Building Data"
    end_marker = "Materials / Layers / Constructions"

    start_index = None
    end_index = None
    for i, line in enumerate(data):
        if start_marker in line:
            start_index = i + 1
        if end_marker in line:
            end_index = i - 1
            break

    if start_index is None or end_index is None:
        raise ValueError("Could not find the specified markers in the file.")

    section = data[start_index:end_index]

    azimuth_found = False
    holidays_index = None

    for i, line in enumerate(section):
        if "AZIMUTH" in line:
            section[i] = f'   AZIMUTH          = {orient}\n'
            azimuth_found = True
        if "HOLIDAYS" in line:
            holidays_index = i

    if not azimuth_found and holidays_index is not None:
        section.insert(holidays_index, f'   AZIMUTH          = {orient}\n')

    data[start_index:end_index] = section

    return data