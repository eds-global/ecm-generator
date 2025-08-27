import pandas as pd
import re

def update_Material_Layers_Construction(inp_content, row_num):
    # Load all necessary sheets
    matData = pd.read_excel('database/AllData.xlsx', sheet_name="Wall")
    material_db = pd.read_excel('database/AllData.xlsx', sheet_name="Material_DB_IP")
    wall_new = pd.read_excel('database/AllData.xlsx', sheet_name="Wall_New")

    if row_num >= len(matData):
        raise IndexError(f"Row index {row_num} is out of bounds for matData.")

    row = matData.iloc[row_num]

    # Extract markers
    start_marker = "Materials / Layers / Constructions"
    end_marker1 = "= LAYERS"
    end_marker2 = "= CONSTRUCTION"

    start_index = inp_content.find(start_marker)
    end_index1 = inp_content.find(end_marker1, start_index)
    end_index2 = inp_content.find(end_marker2, start_index)
    end_index1 = inp_content.rfind('\n', start_index, end_index1)
    end_index2 = inp_content.rfind('\n', start_index, end_index2)

    # Gather material names
    material_names = [row[f'Mat_{i}'] for i in range(1, 5) if pd.notna(row[f'Mat_{i}'])]

    outputMat = ""
    outputLayer = ""
    outputCons = ""

    unique_materials = list(dict.fromkeys(material_names))  # remove duplicates while preserving order
    thicknesses_list = []

    for mat in material_names:
        # MATERIAL section from Material_DB_IP
        matched = material_db[material_db['Code'] == mat]
        if not matched.empty:
            mat_row = matched.iloc[0]
            outputMat += f'"{mat}"  = MATERIAL\n'
            outputMat += f'  TYPE           = PROPERTIES\n'
            outputMat += f'  THICKNESS      = {mat_row["DefaultThick(ft)"]}\n'
            outputMat += f'  CONDUCTIVITY   = {mat_row["Conductivity(BTU/(hr·ft·°F))"]:.2f}\n'
            outputMat += f'  DENSITY        = {mat_row["Density(lb/ft³)"]:.2f}\n'
            outputMat += f'  SPECIFIC-HEAT  = {mat_row["Sp_Heat(BTU/lb·°F)"]:.2f}\n'
            outputMat += f'  ..\n'
        else:
            print(f"WARNING: Material '{mat}' not found in Material_DB_IP.")

        # LAYER thickness from Wall_New based on Material + Construction match
        wall_new_match = wall_new[
            (wall_new['Material'] == mat) &
            (wall_new['Code'] == row['Code'])
        ]
        if not wall_new_match.empty:
            thickness_ft = wall_new_match.iloc[0]['Thickness(ft)']
            thicknesses_list.append(round(thickness_ft, 2))
        else:
            print(f"WARNING: Thickness for material '{mat}' not found in Wall_New for construction '{row['Code']}'")
            thicknesses_list.append(0.0)

    # Create LAYERS block
    mat_string = ', '.join([f'"{m}"' for m in material_names])
    thick_string = ', '.join([str(t) for t in thicknesses_list])

    outputLayer += f'"{row["Code"]}_Lyr"  = LAYERS\n'
    outputLayer += f'  MATERIAL       = ({mat_string})\n'
    outputLayer += f'  THICKNESS      = ({thick_string})\n'
    outputLayer += f'  ..\n\n'

    # Create CONSTRUCTION block
    outputCons += f'"{row["Code"]}"  = CONSTRUCTION\n'
    outputCons += f'  TYPE           = LAYERS\n'
    outputCons += f'  LAYERS         = "{row["Code"]}_Lyr"\n'
    outputCons += f'  ..\n'

    # Insert the new content into the original inp
    updated_inp_content = (
        inp_content[:end_index1] + outputMat + inp_content[end_index1:end_index2] +
        outputLayer + outputCons + inp_content[end_index2:]
    )

    # Update EXTERIOR-WALL constructions
    updated_inp_content = re.sub(
        r'(".*?")\s*=\s*EXTERIOR-WALL\s*CONSTRUCTION\s*=\s*"[^"]+"',
        lambda m: f'{m.group(1)} = EXTERIOR-WALL\n   CONSTRUCTION     = "{row["Code"]}"',
        updated_inp_content
    )

    return updated_inp_content


# import pandas as pd
# import os
# import re

# def update_Material_Layers_Construction(inp_content, row_num):
#     # with open(name, 'r') as file:
#     #     inp_content = file.read()

#     # Extracting relevant section between "Materials / Layers / Constructions" and "= LAYERS"
#     start_marker = "Materials / Layers / Constructions"
#     end_marker1 = "= LAYERS"
#     end_marker2 = "= CONSTRUCTION"

#     start_index = inp_content.find(start_marker)
#     end_index1 = inp_content.find(end_marker1, start_index)
#     end_index2 = inp_content.find(end_marker2, start_index)
#     end_index1 = inp_content.rfind('\n', start_index, end_index1)
#     end_index2 = inp_content.rfind('\n', start_index, end_index2)

#     relevant_section1 = inp_content[start_index:end_index1]
#     relevant_section2 = inp_content[start_index:end_index2]

#     matData = pd.read_excel('database/AllData.xlsx', sheet_name="Wall")

#     # **Fix: Fetch row from DataFrame using row_num**
#     if row_num >= len(matData):  # Ensure row_num is within range
#         print(f"ERROR: row_num={row_num} is out of bounds for matData with {len(matData)} rows.")
#         raise IndexError(f"Row index {row_num} is out of bounds for matData with {len(matData)} rows.")

#     row = matData.iloc[row_num]  # Convert row number to actual row data
#     # Convert row number to actual row data

#     outputMat = ""
#     outputLayer = ""
#     outputCons = ""

#     # Extracting relevant information for the current row
#     material_names = [row[f'Mat_{i}'] for i in range(1, 5)]
#     material_conduct = [row[f'Mat_{i}_conductivity'] for i in range(1, 5)]
#     material_density = [row[f'Mat_{i}_Density'] for i in range(1, 5)]
#     material_spc_heat = [row[f'Mat_{i}_Sp_Heat'] for i in range(1, 5)]
#     thicknesses = row['DefaultThick']

#     unique_materials = set(material_names)

#     for material_name in unique_materials:
#         idx = material_names.index(material_name)
#         outputMat += f'"{material_name}"  = MATERIAL\n'
#         outputMat += f'  TYPE           = PROPERTIES\n'
#         outputMat += f'  THICKNESS      = {thicknesses}\n'
#         outputMat += f'  CONDUCTIVITY   = {material_conduct[idx]:.2f}\n'
#         outputMat += f'  DENSITY        = {material_density[idx]:.2f}\n'
#         outputMat += f'  SPECIFIC-HEAT  = {material_spc_heat[idx]:.2f}\n'
#         outputMat += f'  ..\n'

#     material_str = ', '.join([f'"{m}"' for m in material_names])
#     thicknesses = [round(row[f'Mat_{i}_thickness'], 2) for i in range(1, 4)]
#     outputLayer += f'"{row["Code"]}_Lyr"  = LAYERS\n'
#     outputLayer += f'  MATERIAL       = ({material_str})\n'
#     outputLayer += f'  THICKNESS      = ({", ".join(map(str, thicknesses))})\n'
#     outputLayer += f'  ..\n\n'

#     outputCons += f'"{row["Code"]}"  = CONSTRUCTION\n'
#     outputCons += f'  TYPE           = LAYERS\n'
#     outputCons += f'  LAYERS         = "{row["Code"]}_Lyr"\n'
#     outputCons += f'  ..\n'

#     updated_inp_content = (
#         inp_content[:end_index1] + outputMat + inp_content[end_index1:end_index2] + outputLayer + outputCons + inp_content[end_index2:]
#     )

#     # Replacing construction values in EXTERIOR-WALL sections
#     updated_inp_content = re.sub(
#         r'(".*?")\s*=\s*EXTERIOR-WALL\s*'
#         r'CONSTRUCTION\s*=\s*"[^"]+"', 
#         lambda match: f'{match.group(1)} = EXTERIOR-WALL\n   CONSTRUCTION     = "{row["Code"]}"',
#         updated_inp_content
#     )
    
#     return updated_inp_content