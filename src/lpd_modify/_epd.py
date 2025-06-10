def getEPD(data, epd):
    if epd is None:
        return data

    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    # Join the list into a single string for finding the start and end
    full_text = ''.join(data)

    start_index = full_text.find(start_marker)
    end_index = full_text.find(end_marker)

    if start_index == -1 or end_index == -1:
        raise ValueError("Could not find the specified markers in the file.")

    # Now map those positions back to line indexes
    line_start_index = None
    line_end_index = None
    char_count = 0
    for idx, line in enumerate(data):
        if line_start_index is None and start_marker in line:
            line_start_index = idx
        if line_end_index is None and end_marker in line:
            line_end_index = idx
        if line_start_index is not None and line_end_index is not None:
            break

    # Replace the EQUIPMENT-W/AREA values
    for i in range(line_start_index, line_end_index + 1):
        if "EQUIPMENT-W/AREA" in data[i]:
            data[i] = f"   EQUIPMENT-W/AREA  = ( {epd} )\n"

    return data
