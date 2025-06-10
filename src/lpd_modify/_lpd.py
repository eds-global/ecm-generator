def getLPD(data, lpd):
    if lpd is None:
        return data

    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    # Step 1: Find the line numbers where the markers exist
    start_index = None
    end_index = None

    for idx, line in enumerate(data):
        if start_marker in line:
            start_index = idx
        elif end_marker in line:
            end_index = idx
        if start_index is not None and end_index is not None:
            break

    if start_index is None or end_index is None:
        raise ValueError("Could not find the specified markers in the file.")

    # Step 2: Modify LIGHTING-W/AREA in the marked section
    for i in range(start_index, end_index + 1):
        if "LIGHTING-W/AREA" in data[i]:
            data[i] = f"   LIGHTING-W/AREA  = ( {lpd} )\n"

    return data
