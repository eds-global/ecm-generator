def getOrientation(data, orient):
    if orient is None:
        return data

    # Convert list of lines to a single string if needed
    if isinstance(data, list):
        data = ''.join(data)

    start_marker = "Site and Building Data"
    end_marker = "Materials / Layers / Constructions"

    start_index = data.find(start_marker)
    end_index = data.find(end_marker)

    if start_index == -1 or end_index == -1:
        raise ValueError("Could not find the specified markers in the file.")

    # Extract the section
    section = data[start_index:end_index].splitlines()

    azimuth_found = False
    holidays_index = None

    # Modify the section
    for i, line in enumerate(section):
        if "AZIMUTH" in line:
            section[i] = f'   AZIMUTH          = {orient}'
            azimuth_found = True
        if "HOLIDAYS" in line:
            holidays_index = i

    if not azimuth_found and holidays_index is not None:
        section.insert(holidays_index, f'   AZIMUTH          = {orient}')

    # Join section back into a string
    updated_section = "\n".join(section)

    # Reconstruct full content
    updated_data = data[:start_index] + updated_section + data[end_index:]

    return updated_data
