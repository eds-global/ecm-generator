import os
import pandas as pd
import re
import streamlit as st

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
    elif row_num > 0 and row_num < 13:
        # -------------------------------
        # R-VALUE MAP
        # -------------------------------
        excel_path = "database/AllData.xlsx"
        # Read Excel
        df = pd.read_excel(excel_path, sheet_name="Wall_New")
        # Skip first 2 rows
        df = df.iloc[3:].reset_index(drop=True)
        # Build map: 1 -> thickness, 2 -> thickness, ...
        rvalue_map = {
            i + 1: float(row["Thicknes(ft)"])
            for i, row in df.iterrows()
        }
        rvalue = rvalue_map[row_num]
        # st.write(rvalue)

        # -------------------------------
        # MATERIAL BLOCKS
        # -------------------------------
        cp_material = (
            '"CP"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            '  THICKNESS       = 0.0656\n'
            '  CONDUCTIVITY    = 0.416\n'
            '  DENSITY         = 110\n'
            '  SPECIFIC-HEAT   = 0.2\n'
            '  ..\n\n'
        )

        xps_material = (
            '"XPS"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            f'  THICKNESS       = {rvalue:.3f}\n'
            '  CONDUCTIVITY    = 0.0161\n'
            '  DENSITY         = 2.18\n'
            '  SPECIFIC-HEAT   = 0.29\n'
            '  ..\n\n'
        )

        brick_material = (
            '"Brick"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            f'  THICKNESS       = 0.7545\n'
            '  CONDUCTIVITY    = 0.5662\n'
            '  DENSITY         = 120\n'
            '  SPECIFIC-HEAT   = 0.2\n'
            '  ..\n\n'
        )

        # cp2_material = cp_material.replace("CP1_RESISTANCE", "CP2_RESISTANCE")
        material_block = cp_material + xps_material + brick_material

        # -------------------------------
        # LAYERS BLOCK
        # -------------------------------
        layer_block = (
            '"AWESIM_WALL_LYR"  = LAYERS\n'
            '  MATERIAL       = ("CP", "XPS", "Brick", "CP")\n'
            '  THICKNESS      = (0.0656, {:.3f}, 0.754, 0.0656)\n'
            '  ..\n\n'
        ).format(rvalue)

        # -------------------------------
        # CONSTRUCTION BLOCK
        # -------------------------------
        construction_block = (
            '"AWESIM_WALL_CONST"  = CONSTRUCTION\n'
            '  TYPE           = LAYERS\n'
            '  LAYERS         = "AWESIM_WALL_LYR"\n'
            '  ..\n\n'
        )

        # -------------------------------
        # INSERT LOCATION
        # -------------------------------
        start_marker = "Materials / Layers / Constructions"
        end_marker = "= LAYERS"

        start_index = inp_content.find(start_marker)
        end_index = inp_content.find(end_marker, start_index)
        end_index = inp_content.rfind("\n", start_index, end_index)

        updated_inp = (
            inp_content[:end_index]
            + "\n\n"
            + material_block
            + layer_block
            + construction_block
            + inp_content[end_index:]
        )

        # Replace only EXTERIOR-WALL blocks where LOCATION != TOP
        def replace_exterior_wall_construction(match):
            block = match.group(0)

            # Skip blocks where LOCATION = TOP (Roof)
            if not re.search(r'\bLOCATION\s*=\s*(TOP|BOTTOM)\b', block, re.IGNORECASE):
                block = re.sub(
                    r'(CONSTRUCTION\s*=\s*)"[^"]+"',
                    r'\1"AWESIM_WALL_CONST"',
                    block,
                    flags=re.IGNORECASE
                )

            return block

        updated_inp = re.sub(
            r'(?im)^\s*(".*?")\s*=\s*EXTERIOR-WALL\b[\s\S]*?^\s*\.\.',
            replace_exterior_wall_construction,
            updated_inp
        )

        return updated_inp
    
    else:
        excel_path = "database/AllData.xlsx"
        df1 = pd.read_excel(excel_path, sheet_name="Wall_New")
        # Excel rows start at 18 for row_num = 13
        df_else = df1.iloc[19:].reset_index(drop=True)
        # Map row_num → df_else index
        local_index = 20 + row_num - 13
        rvalue = float(df1.loc[local_index, "Thicknes(ft)"])

        # -------------------------------
        # MATERIAL BLOCKS
        # -------------------------------
        cb_material = (
            '"CB"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            '  THICKNESS       = 0.06562\n'
            '  CONDUCTIVITY    = 0.329346\n'
            '  DENSITY         = 118.617\n'
            '  SPECIFIC-HEAT   = 0.2388\n'
            '  ..\n\n'
        )

        xps_material = (
            '"XPS"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            f'  THICKNESS       = {rvalue:.4f}\n'
            '  CONDUCTIVITY    = 0.0161\n'
            '  DENSITY         = 2.18\n'
            '  SPECIFIC-HEAT   = 0.29\n'
            '  ..\n\n'
        )

        gyp_material = (
            '"GYP"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            f'  THICKNESS       = 0.04921\n'
            '  CONDUCTIVITY    = 0.092448\n'
            '  DENSITY         = 39.9552\n'
            '  SPECIFIC-HEAT   = 0.27462\n'
            '  ..\n\n'
        )

        material_block = cb_material + xps_material + gyp_material

        # -------------------------------
        # LAYERS BLOCK
        # -------------------------------
        layer_block = (
            '"AWESIM_WALL_LYR"  = LAYERS\n'
            '  MATERIAL       = ("CB", "XPS", "GYP")\n'
            '  THICKNESS      = (0.06562, {:.4f}, 0.04921)\n'
            '  ..\n\n'
        ).format(rvalue)

        # -------------------------------
        # CONSTRUCTION BLOCK
        # -------------------------------
        construction_block = (
            '"AWESIM_WALL_CONST"  = CONSTRUCTION\n'
            '  TYPE           = LAYERS\n'
            '  LAYERS         = "AWESIM_WALL_LYR"\n'
            '  ..\n\n'
        )

        # -------------------------------
        # INSERT LOCATION
        # -------------------------------
        start_marker = "Materials / Layers / Constructions"
        end_marker = "= LAYERS"

        start_index = inp_content.find(start_marker)
        end_index = inp_content.find(end_marker, start_index)
        end_index = inp_content.rfind("\n", start_index, end_index)

        updated_inp = (
            inp_content[:end_index]
            + "\n\n"
            + material_block
            + layer_block
            + construction_block
            + inp_content[end_index:]
        )

        # Replace only EXTERIOR-WALL blocks where LOCATION != TOP
        def replace_exterior_wall_construction(match):
            block = match.group(0)

            # Skip blocks where LOCATION = TOP (Roof)
            if not re.search(r'\bLOCATION\s*=\s*(TOP|BOTTOM)\b', block, re.IGNORECASE):
                block = re.sub(
                    r'(CONSTRUCTION\s*=\s*)"[^"]+"',
                    r'\1"AWESIM_WALL_CONST"',
                    block,
                    flags=re.IGNORECASE
                )

            return block

        updated_inp = re.sub(
            r'(?im)^\s*(".*?")\s*=\s*EXTERIOR-WALL\b[\s\S]*?^\s*\.\.',
            replace_exterior_wall_construction,
            updated_inp
        )

        return updated_inp

def fix_roofs(inp_content, row_num):
    if isinstance(inp_content, list):
        inp_content = "".join(inp_content)

    if row_num == 0:
        return inp_content
    elif row_num > 0 and row_num < 13:
        # -------------------------------
        # R-VALUE MAP
        # -------------------------------
        excel_path = "database/AllData.xlsx"
        # Read Excel
        df = pd.read_excel(excel_path, sheet_name="Roof_New")
        # Skip first 2 rows
        df = df.iloc[4:].reset_index(drop=True)
        # Build map: 1 -> thickness, 2 -> thickness, ...
        rvalue_map = {
            i + 1: float(row["Thicknes(ft)"])
            for i, row in df.iterrows()
        }
        rvalue = rvalue_map[row_num]
        # st.write(rvalue)

        # -------------------------------
        # MATERIAL BLOCKS
        # -------------------------------
        cp_material = (
            '"CP"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            '  THICKNESS       = 0.0656\n'
            '  CONDUCTIVITY    = 0.416\n'
            '  DENSITY         = 110\n'
            '  SPECIFIC-HEAT   = 0.2\n'
            '  ..\n\n'
        )

        xps_material = (
            '"XPS"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            f'  THICKNESS       = {rvalue:.3f}\n'
            '  CONDUCTIVITY    = 0.0161\n'
            '  DENSITY         = 2.18\n'
            '  SPECIFIC-HEAT   = 0.29\n'
            '  ..\n\n'
        )

        brick_material = (
            '"Brick"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            f'  THICKNESS       = 0.25\n'
            '  CONDUCTIVITY    = 0.566\n'
            '  DENSITY         = 120\n'
            '  SPECIFIC-HEAT   = 0.2\n'
            '  ..\n\n'
        )

        concrete_material = (
            '"Concrete"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            f'  THICKNESS      = 0.4921\n'
            '  CONDUCTIVITY    = 0.9129\n'
            '  DENSITY         = 142.84\n'
            '  SPECIFIC-HEAT   = 0.21\n'
            '  ..\n\n'
        )

        # cp2_material = cp_material.replace("CP1_RESISTANCE", "CP2_RESISTANCE")
        material_block = cp_material + xps_material + brick_material + concrete_material

        # -------------------------------
        # LAYERS BLOCK
        # -------------------------------
        layer_block = (
            '"AWESIM_ROOF_LYR"  = LAYERS\n'
            '  MATERIAL       = ("CP", "XPS", "Brick", "Concrete", "CP")\n'
            '  THICKNESS      = (0.0656, {:.3f}, 0.25, 0.492, 0.0656)\n'
            '  ..\n\n'
        ).format(rvalue)

        # -------------------------------
        # CONSTRUCTION BLOCK
        # -------------------------------
        construction_block = (
            '"AWESIM_ROOF_CONST"  = CONSTRUCTION\n'
            '  TYPE           = LAYERS\n'
            '  LAYERS         = "AWESIM_ROOF_LYR"\n'
            '  ..\n\n'
        )

        # -------------------------------
        # INSERT LOCATION
        # -------------------------------
        start_marker = "Materials / Layers / Constructions"
        end_marker = "= LAYERS"

        start_index = inp_content.find(start_marker)
        end_index = inp_content.find(end_marker, start_index)
        end_index = inp_content.rfind("\n", start_index, end_index)
        

        updated_inp = (
            inp_content[:end_index]
            + "\n\n"
            + material_block
            + layer_block
            + construction_block
            + inp_content[end_index:]
        )

        # Replace only EXTERIOR-WALL blocks where LOCATION = TOP (Roof)
        def replace_exterior_wall_construction(match):
            block = match.group(0)

            # Ensure LOCATION = TOP exists in the same block
            if re.search(r'\bLOCATION\s*=\s*TOP\b', block, re.IGNORECASE):
                block = re.sub(
                    r'(CONSTRUCTION\s*=\s*)"[^"]+"',
                    rf'\1"AWESIM_ROOF_CONST"',
                    block,
                    flags=re.IGNORECASE
                )

            return block

        updated_inp = re.sub(
            r'(?im)^\s*(".*?")\s*=\s*EXTERIOR-WALL\b[\s\S]*?^\s*\.\.',
            replace_exterior_wall_construction,
            updated_inp
        )

        return updated_inp
    else:
        # -------------------------------
        # R-VALUE MAP (CORRECTED)
        # -------------------------------
        excel_path = "database/AllData.xlsx"
        df1 = pd.read_excel(excel_path, sheet_name="Roof_New")
        # Excel rows start at 18 for row_num = 13
        df_else = df1.iloc[19:].reset_index(drop=True)
        # Map row_num → df_else index
        local_index = 19 + row_num - 13
        rvalue = float(df1.loc[local_index, "Thicknes(ft)"])
        # st.write(local_index)
        # st.write(rvalue)
        # -------------------------------
        # MATERIAL BLOCKS
        # -------------------------------
        metDeck_material = (
            '"MetDeck"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            '  THICKNESS       = 0.021\n'
            '  CONDUCTIVITY    = 26.2000\n'
            '  DENSITY         = 489.00\n'
            '  SPECIFIC-HEAT   = 0.12\n'
            '  ..\n\n'
        )
        xps_material = (
            '"XPS"  = MATERIAL\n'
            '  TYPE            = PROPERTIES\n'
            f'  THICKNESS       = {rvalue:.3f}\n'
            '  CONDUCTIVITY    = 0.0161\n'
            '  DENSITY         = 2.18\n'
            '  SPECIFIC-HEAT   = 0.29\n'
            '  ..\n\n'
        )
        material_block = metDeck_material + xps_material
        # -------------------------------
        # LAYERS BLOCK
        # -------------------------------
        layer_block = (
            '"AWESIM_ROOF_LYR"  = LAYERS\n'
            '  MATERIAL       = ("MetDeck", "XPS")\n'
            '  THICKNESS      = (0.021, {:.3f})\n'
            '  ..\n\n'
        ).format(rvalue)
        # -------------------------------
        # CONSTRUCTION BLOCK
        # -------------------------------
        construction_block = (
            '"AWESIM_ROOF_CONST"  = CONSTRUCTION\n'
            '  TYPE           = LAYERS\n'
            '  LAYERS         = "AWESIM_ROOF_LYR"\n'
            '  ..\n\n'
        )

        # -------------------------------
        # INSERT LOCATION
        # -------------------------------
        start_marker = "Materials / Layers / Constructions"
        end_marker = "= LAYERS"

        start_index = inp_content.find(start_marker)
        end_index = inp_content.find(end_marker, start_index)
        end_index = inp_content.rfind("\n", start_index, end_index)
        

        updated_inp = (
            inp_content[:end_index]
            + "\n\n"
            + material_block
            + layer_block
            + construction_block
            + inp_content[end_index:]
        )

        # -------------------------------
        # UPDATE EXTERIOR-WALL REFERENCES
        # -------------------------------
        # updated_inp = re.sub(
        #     r'(".*?")\s*=\s*EXTERIOR-WALL\s*CONSTRUCTION\s*=\s*"[^"]+"',
        #     lambda m: (
        #         f'{m.group(1)} = EXTERIOR-WALL\n'
        #         f'   CONSTRUCTION     = "WALL_R_CONST"'
        #     ),
        #     updated_inp
        # )

        # return updated_inp


        # Replace only EXTERIOR-WALL blocks where LOCATION = TOP (Roof)
        def replace_exterior_wall_construction(match):
            block = match.group(0)

            # Ensure LOCATION = TOP exists in the same block
            if re.search(r'\bLOCATION\s*=\s*TOP\b', block, re.IGNORECASE):
                block = re.sub(
                    r'(CONSTRUCTION\s*=\s*)"[^"]+"',
                    rf'\1"AWESIM_ROOF_CONST"',
                    block,
                    flags=re.IGNORECASE
                )

            return block

        updated_inp = re.sub(
            r'(?im)^\s*(".*?")\s*=\s*EXTERIOR-WALL\b[\s\S]*?^\s*\.\.',
            replace_exterior_wall_construction,
            updated_inp
        )

        return updated_inp

# import os
# import pandas as pd
# import re
# import streamlit as st

# def count_exterior_walls(inp_file):
#     start_marker = "Floors / Spaces / Walls / Windows / Doors"
#     end_marker = "Electric & Fuel Meters"
    
#     start_index, end_index = None, None
    
#     # find marker positions
#     for i, line in enumerate(inp_file):
#         if start_marker in line:
#             start_index = i
#         if end_marker in line and start_index is not None:
#             end_index = i
#             break
    
#     if start_index is None or end_index is None:
#         raise ValueError("Could not find section markers in INP file.")
    
#     # slice between markers
#     section = inp_file[start_index:end_index]
    
#     # regex to find any object definition like "Wall_xxx" = EXTERIOR-WALL
#     pattern = re.compile(r'^\s*".*"\s*=\s*EXTERIOR-WALL')
#     count = sum(1 for line in section if pattern.search(line))
    
#     return count

# def wrap_line_at_comma(line, width=75, indent="         "):
#     """
#     Wraps a line at the last comma before `width`.
#     If no comma found before width, keeps line as is.
#     """
#     if len(line) <= width:
#         return line

#     parts = []
#     current = line
#     while len(current) > width:
#         # find last comma before width
#         break_pos = current.rfind(",", 0, width)
#         if break_pos == -1:
#             break  # no comma, stop wrapping

#         # include the comma at end of line
#         parts.append(current[:break_pos+1].rstrip())
#         # indent continuation lines
#         current = indent + current[break_pos+1:].lstrip()

#     parts.append(current)
#     return "\n".join(parts)

# def fix_walls(inp_content, row_num):
#     if isinstance(inp_content, list):
#         inp_content = "".join(inp_content)

#     if row_num == 0:
#         return inp_content

#     # -------------------------------
#     # R-VALUE MAP
#     # -------------------------------
#     excel_path = r"D:\EDS\260108AWESIM\database\AllData.xlsx"
#     # Read Excel
#     df = pd.read_excel(excel_path, sheet_name="Wall_New")
#     # Skip first 2 rows
#     df = df.iloc[3:].reset_index(drop=True)
#     # Build map: 1 -> thickness, 2 -> thickness, ...
#     rvalue_map = {
#         i + 1: float(row["Thicknes(ft)"])
#         for i, row in df.iterrows()
#     }
#     rvalue = rvalue_map[row_num]
#     # st.write(rvalue)

#     # -------------------------------
#     # MATERIAL BLOCKS
#     # -------------------------------
#     cp_material = (
#         '"CP"  = MATERIAL\n'
#         '  TYPE            = PROPERTIES\n'
#         '  THICKNESS       = 0.065\n'
#         '  CONDUCTIVITY    = 0.416\n'
#         '  DENSITY         = 110\n'
#         '  SPECIFIC-HEAT   = 0.2\n'
#         '  ..\n\n'
#     )

#     brick_material = (
#         '"Brick"  = MATERIAL\n'
#         '  TYPE            = PROPERTIES\n'
#         f'  THICKNESS       = 0.754\n'
#         '  CONDUCTIVITY    = 0.566\n'
#         '  DENSITY         = 120\n'
#         '  SPECIFIC-HEAT   = 0.2\n'
#         '  ..\n\n'
#     )

#     xps_material = (
#         '"XPS"  = MATERIAL\n'
#         '  TYPE            = PROPERTIES\n'
#         f'  THICKNESS       = {rvalue:.3f}\n'
#         '  CONDUCTIVITY    = 0.016\n'
#         '  DENSITY         = 2.18\n'
#         '  SPECIFIC-HEAT   = 0.29\n'
#         '  ..\n\n'
#     )

#     # cp2_material = cp_material.replace("CP1_RESISTANCE", "CP2_RESISTANCE")

#     material_block = cp_material + brick_material + xps_material

#     # -------------------------------
#     # LAYERS BLOCK
#     # -------------------------------
#     layer_block = (
#         '"AWESIM_WALL_LYR"  = LAYERS\n'
#         '  MATERIAL       = ("CP", "Brick", "XPS", "CP")\n'
#         '  THICKNESS      = (0.065, 0.754, {:.3f}, 0.065)\n'
#         '  ..\n\n'
#     ).format(rvalue)

#     # -------------------------------
#     # CONSTRUCTION BLOCK
#     # -------------------------------
#     construction_block = (
#         '"AWESIM_WALL_CONST"  = CONSTRUCTION\n'
#         '  TYPE           = LAYERS\n'
#         '  LAYERS         = "AWESIM_WALL_LYR"\n'
#         '  ..\n\n'
#     )

#     # -------------------------------
#     # INSERT LOCATION
#     # -------------------------------
#     start_marker = "Materials / Layers / Constructions"
#     end_marker = "= LAYERS"

#     start_index = inp_content.find(start_marker)
#     end_index = inp_content.find(end_marker, start_index)
#     end_index = inp_content.rfind("\n", start_index, end_index)

#     updated_inp = (
#         inp_content[:end_index]
#         + "\n\n"
#         + material_block
#         + layer_block
#         + construction_block
#         + inp_content[end_index:]
#     )

#     # -------------------------------
#     # UPDATE EXTERIOR-WALL REFERENCES
#     # -------------------------------
#     # updated_inp = re.sub(
#     #     r'(".*?")\s*=\s*EXTERIOR-WALL\s*CONSTRUCTION\s*=\s*"[^"]+"',
#     #     lambda m: (
#     #         f'{m.group(1)} = EXTERIOR-WALL\n'
#     #         f'   CONSTRUCTION     = "WALL_CONST"'
#     #     ),
#     #     updated_inp
#     # )

#     # Replace only EXTERIOR-WALL blocks where LOCATION != TOP
#     def replace_exterior_wall_construction(match):
#         block = match.group(0)

#         # Skip blocks where LOCATION = TOP (Roof)
#         if not re.search(r'\bLOCATION\s*=\s*(TOP|BOTTOM)\b', block, re.IGNORECASE):
#             block = re.sub(
#                 r'(CONSTRUCTION\s*=\s*)"[^"]+"',
#                 r'\1"AWESIM_WALL_CONST"',
#                 block,
#                 flags=re.IGNORECASE
#             )

#         return block

#     updated_inp = re.sub(
#         r'(?im)^\s*(".*?")\s*=\s*EXTERIOR-WALL\b[\s\S]*?^\s*\.\.',
#         replace_exterior_wall_construction,
#         updated_inp
#     )


#     return updated_inp

# def fix_roofs(inp_content, row_num):
#     if isinstance(inp_content, list):
#         inp_content = "".join(inp_content)

#     if row_num == 0:
#         return inp_content

#     # -------------------------------
#     # R-VALUE MAP
#     # -------------------------------
#     excel_path = r"D:\EDS\260108AWESIM\database\AllData.xlsx"
#     # Read Excel
#     df = pd.read_excel(excel_path, sheet_name="Roof_New")
#     # Skip first 2 rows
#     df = df.iloc[4:].reset_index(drop=True)
#     # Build map: 1 -> thickness, 2 -> thickness, ...
#     rvalue_map = {
#         i + 1: float(row["Thicknes(ft)"])
#         for i, row in df.iterrows()
#     }
#     rvalue = rvalue_map[row_num]
#     # st.write(rvalue)

#     # -------------------------------
#     # MATERIAL BLOCKS
#     # -------------------------------
#     cp_material = (
#         '"CP"  = MATERIAL\n'
#         '  TYPE            = PROPERTIES\n'
#         '  THICKNESS       = 0.065\n'
#         '  CONDUCTIVITY    = 0.416\n'
#         '  DENSITY         = 110\n'
#         '  SPECIFIC-HEAT   = 0.2\n'
#         '  ..\n\n'
#     )

#     brick_material = (
#         '"Brick"  = MATERIAL\n'
#         '  TYPE            = PROPERTIES\n'
#         f'  THICKNESS       = 0.250\n'
#         '  CONDUCTIVITY    = 0.566\n'
#         '  DENSITY         = 120\n'
#         '  SPECIFIC-HEAT   = 0.2\n'
#         '  ..\n\n'
#     )

#     xps_material = (
#         '"XPS"  = MATERIAL\n'
#         '  TYPE            = PROPERTIES\n'
#         f'  THICKNESS       = {rvalue:.3f}\n'
#         '  CONDUCTIVITY    = 0.016\n'
#         '  DENSITY         = 2.18\n'
#         '  SPECIFIC-HEAT   = 0.29\n'
#         '  ..\n\n'
#     )

#     concrete_material = (
#         '"Concrete"  = MATERIAL\n'
#         '  TYPE            = PROPERTIES\n'
#         f'  THICKNESS      = 0.492\n'
#         '  CONDUCTIVITY    = 0.913\n'
#         '  DENSITY         = 142.84\n'
#         '  SPECIFIC-HEAT   = 0.21\n'
#         '  ..\n\n'
#     )

#     # cp2_material = cp_material.replace("CP1_RESISTANCE", "CP2_RESISTANCE")

#     material_block = cp_material + brick_material + xps_material + concrete_material

#     # -------------------------------
#     # LAYERS BLOCK
#     # -------------------------------
#     layer_block = (
#         '"AWESIM_ROOF_LYR"  = LAYERS\n'
#         '  MATERIAL       = ("CP", "Brick", "XPS", "Concrete", "CP")\n'
#         '  THICKNESS      = (0.065, 0.754, {:.3f}, 0.492, 0.065)\n'
#         '  ..\n\n'
#     ).format(rvalue)

#     # -------------------------------
#     # CONSTRUCTION BLOCK
#     # -------------------------------
#     construction_block = (
#         '"AWESIM_ROOF_CONST"  = CONSTRUCTION\n'
#         '  TYPE           = LAYERS\n'
#         '  LAYERS         = "AWESIM_ROOF_LYR"\n'
#         '  ..\n\n'
#     )

#     # -------------------------------
#     # INSERT LOCATION
#     # -------------------------------
#     start_marker = "Materials / Layers / Constructions"
#     end_marker = "= LAYERS"

#     start_index = inp_content.find(start_marker)
#     end_index = inp_content.find(end_marker, start_index)
#     end_index = inp_content.rfind("\n", start_index, end_index)
    

#     updated_inp = (
#         inp_content[:end_index]
#         + "\n\n"
#         + material_block
#         + layer_block
#         + construction_block
#         + inp_content[end_index:]
#     )

#     # -------------------------------
#     # UPDATE EXTERIOR-WALL REFERENCES
#     # -------------------------------
#     # updated_inp = re.sub(
#     #     r'(".*?")\s*=\s*EXTERIOR-WALL\s*CONSTRUCTION\s*=\s*"[^"]+"',
#     #     lambda m: (
#     #         f'{m.group(1)} = EXTERIOR-WALL\n'
#     #         f'   CONSTRUCTION     = "WALL_R_CONST"'
#     #     ),
#     #     updated_inp
#     # )

#     # return updated_inp


#     # Replace only EXTERIOR-WALL blocks where LOCATION = TOP (Roof)
#     def replace_exterior_wall_construction(match):
#         block = match.group(0)

#         # Ensure LOCATION = TOP exists in the same block
#         if re.search(r'\bLOCATION\s*=\s*TOP\b', block, re.IGNORECASE):
#             block = re.sub(
#                 r'(CONSTRUCTION\s*=\s*)"[^"]+"',
#                 rf'\1"AWESIM_ROOF_CONST"',
#                 block,
#                 flags=re.IGNORECASE
#             )

#         return block

#     updated_inp = re.sub(
#         r'(?im)^\s*(".*?")\s*=\s*EXTERIOR-WALL\b[\s\S]*?^\s*\.\.',
#         replace_exterior_wall_construction,
#         updated_inp
#     )

#     return updated_inp