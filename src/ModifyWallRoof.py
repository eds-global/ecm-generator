import os
import pandas as pd
import re

def count_exterior_walls(inp_file):
    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"
    
    start_index, end_index = None, None
    
    # find marker positions
    for i, line in enumerate(inp_file):
        if start_marker in line:
            start_index = i
        if end_marker in line and start_index is not None:
            end_index = i
            break
    
    if start_index is None or end_index is None:
        raise ValueError("Could not find section markers in INP file.")
    
    # slice between markers
    section = inp_file[start_index:end_index]
    
    # regex to find any object definition like "Wall_xxx" = EXTERIOR-WALL
    pattern = re.compile(r'^\s*".*"\s*=\s*EXTERIOR-WALL')
    count = sum(1 for line in section if pattern.search(line))
    
    return count

import re

def wrap_line_at_comma(line, width=75, indent="         "):
    """
    Wraps a line at the last comma before `width`.
    If no comma found before width, keeps line as is.
    """
    if len(line) <= width:
        return line

    parts = []
    current = line
    while len(current) > width:
        # find last comma before width
        break_pos = current.rfind(",", 0, width)
        if break_pos == -1:
            break  # no comma, stop wrapping

        # include the comma at end of line
        parts.append(current[:break_pos+1].rstrip())
        # indent continuation lines
        current = indent + current[break_pos+1:].lstrip()

    parts.append(current)
    return "\n".join(parts)


def fix_walls(inp_content, row_num):
    if isinstance(inp_content, list):
        inp_content = "".join(inp_content)

    if row_num == 0:
        return inp_content

    # --- RValue map ---
    rvalue_map = {1: 2.5, 2: 5.0, 3: 7.5, 4: 10, 5: 12.5, 6: 15, 7: 17.5, 8: 20, 9: 22.5, 10: 25, 11: 27.5, 12: 30}
    rvalue = rvalue_map.get(row_num, 0.0)

    # --- Define new MATERIAL block ---
    mat_name = "ECM_RESISTANCE"
    material_block = f'"{mat_name}"  = MATERIAL\n'
    material_block += '  TYPE            = RESISTANCE\n'
    material_block += f'  RESISTANCE      = {rvalue:.2f}\n'
    material_block += '  ..\n'

    lines = inp_content.splitlines(True)  # keep line endings
    new_lines = []
    inside_layers = False
    inside_wall = False
    mat_done = False

    for i, line in enumerate(lines):
        # Insert ECM_PUF MATERIAL definition block before the first = LAYERS
        if not mat_done and "= LAYERS" in line:
            new_lines.append(material_block)
            mat_done = True

        # Detect start of LAYERS
        if "= LAYERS" in line and ("UW" not in line and "UF" not in line and "Roof" not in line):
            inside_layers = True
            inside_wall = True

        # If inside LAYERS and line has MATERIAL, insert ECM_PUF at 2nd pos
        if inside_layers and "MATERIAL" in line:
            line = re.sub(
                r'MATERIAL\s*=\s*\(\s*"([^"]+)"',
                r'MATERIAL         = ( "\1", "ECM_RESISTANCE"',
                line,
                count=1
            )

        # If inside LAYERS and line has THICKNESS, insert &D at 2nd pos
        if inside_layers and "THICKNESS" in line:
            line = re.sub(
                r'THICKNESS\s*=\s*\(\s*([^,]+)',
                r'THICKNESS        = ( \1, &D',
                line,
                count=1
            )

        # Wrap long MATERIAL/THICKNESS lines at 75 cols
        if inside_layers and inside_wall and ("MATERIAL" in line or "THICKNESS" in line):
            line = wrap_line_at_comma(line, width=75, indent="         ")

        # End of a LAYERS block
        if inside_layers and line.strip().startswith(".."):
            inside_layers = False

        new_lines.append(line)

    updated_inp = "".join(new_lines)
    return updated_inp


def fix_roofs(inp_content, row_num):
    if isinstance(inp_content, list):
        inp_content = "".join(inp_content)

    if row_num == 0:
        return inp_content

    # --- RValue map ---
    rvalue_map = {1: 2.5, 2: 5.0, 3: 7.5, 4: 10, 5: 12.5, 6: 15, 7: 17.5, 8: 20, 9: 22.5, 10: 25, 11: 27.5, 12: 30}
    rvalue = rvalue_map.get(row_num, 0.0)

    # --- Define new MATERIAL block ---
    mat_name = "ECM_RESISTANCE"
    material_block = f'"{mat_name}"  = MATERIAL\n'
    material_block += '  TYPE            = RESISTANCE\n'
    material_block += f'  RESISTANCE      = {rvalue:.2f}\n'
    material_block += '  ..\n'

    lines = inp_content.splitlines(True)  # keep line endings
    new_lines = []
    inside_layers = False
    inside_roof = False
    mat_done = False

    for i, line in enumerate(lines):
        # Insert ECM_PUF MATERIAL definition block before the first = LAYERS
        if not mat_done and "= LAYERS" in line:
            new_lines.append(material_block)
            mat_done = True

        # Detect start of LAYERS
        if "= LAYERS" in line and ("UW" not in line and "UF" not in line and "Wall" not in line):
            inside_layers = True
            inside_roof = True

        # If inside LAYERS and line has MATERIAL, insert ECM_PUF at 2nd pos
        if inside_layers and "MATERIAL" in line:
            line = re.sub(
                r'MATERIAL\s*=\s*\(\s*"([^"]+)"',
                r'MATERIAL         = ( "\1", "ECM_RESISTANCE"',
                line,
                count=1
            )

        # If inside LAYERS and line has THICKNESS, insert &D at 2nd pos
        if inside_layers and "THICKNESS" in line:
            line = re.sub(
                r'THICKNESS\s*=\s*\(\s*([^,]+)',
                r'THICKNESS        = ( \1, &D',
                line,
                count=1
            )

        # Wrap long MATERIAL/THICKNESS lines at 75 cols
        if inside_layers and inside_roof and ("MATERIAL" in line or "THICKNESS" in line):
            line = wrap_line_at_comma(line, width=75, indent="         ")

        # End of a LAYERS block
        if inside_layers and line.strip().startswith(".."):
            inside_layers = False

        new_lines.append(line)

    updated_inp = "".join(new_lines)
    return updated_inp
