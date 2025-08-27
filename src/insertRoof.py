import pandas as pd
import re

def removeDuplicates(inp_content):
    start_marker = "Materials / Layers / Constructions"
    end_marker = "Glass Type Codes"

    inp_data = inp_content.splitlines(keepends=True)  # Keep \n in lines
    start_index, end_index = None, None

    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 4  # Skip header
        if end_marker in line:
            end_index = i - 4
            break

    if start_index is None or end_index is None:
        raise ValueError("Could not find section markers in INP file.")

    # Extract section to deduplicate
    section = inp_data[start_index:end_index+1]
    cleaned_section = []
    seen_materials = set()
    i = 0

    while i < len(section):
        line = section[i]
        material_match = re.match(r'^\s*"([^"]+)"\s*=\s*MATERIAL', line)
        if material_match:
            material_name = material_match.group(1)
            material_block = [line]
            i += 1

            # Collect full material block
            while i < len(section) and section[i].strip() != "..":
                material_block.append(section[i])
                i += 1
            if i < len(section):
                material_block.append(section[i])  # Append ".."
                i += 1

            if material_name not in seen_materials:
                seen_materials.add(material_name)
                cleaned_section.extend(material_block)
            # else skip block (duplicate)
        else:
            cleaned_section.append(line)
            i += 1

    # Merge the full content
    new_inp_data = inp_data[:start_index] + cleaned_section + inp_data[end_index+1:]
    return ''.join(new_inp_data)

def update_Material_Layers_Construction(inp_content, row_num):
    # Load sheets
    matData = pd.read_excel('database/AllData.xlsx', sheet_name="Roof")
    material_db = pd.read_excel('database/AllData.xlsx', sheet_name="Material_DB_IP")
    roof_new = pd.read_excel('database/AllData.xlsx', sheet_name="Roof_New")

    # Find section markers
    start_marker = "Materials / Layers / Constructions"
    end_marker1 = "= LAYERS"
    end_marker2 = "= CONSTRUCTION"

    start_index = inp_content.find(start_marker)
    end_index1 = inp_content.find(end_marker1, start_index)
    end_index2 = inp_content.find(end_marker2, start_index)
    end_index1 = inp_content.rfind('\n', start_index, end_index1)
    end_index2 = inp_content.rfind('\n', start_index, end_index2)

    # Validate row number
    if row_num >= len(matData):
        raise IndexError(f"Row index {row_num} is out of bounds for matData.")
    row = matData.iloc[row_num]

    outputMat = ""
    outputLayer = ""
    outputCons = ""

    # Extract unique materials
    material_names = [row[f'Mat_{i}'] for i in range(1, 5) if pd.notna(row[f'Mat_{i}'])]
    unique_materials = list(dict.fromkeys(material_names))  # preserve order
    thicknesses_list = []

    for mat in material_names:
        # MATERIAL block from Material_DB_IP
        match = material_db[material_db['Code'] == mat]
        if not match.empty:
            mat_row = match.iloc[0]
            outputMat += f'"{mat}"  = MATERIAL\n'
            outputMat += f'  TYPE           = PROPERTIES\n'
            outputMat += f'  THICKNESS      = {mat_row["DefaultThick(ft)"]}\n'
            outputMat += f'  CONDUCTIVITY   = {mat_row["Conductivity(BTU/(hr·ft·°F))"]:.2f}\n'
            outputMat += f'  DENSITY        = {mat_row["Density(lb/ft³)"]:.2f}\n'
            outputMat += f'  SPECIFIC-HEAT  = {mat_row["Sp_Heat(BTU/lb·°F)"]:.2f}\n'
            outputMat += f'  ..\n'
        else:
            print(f"WARNING: Material '{mat}' not found in Material_DB_IP.")

        # LAYERS thickness from Roof_New
        roof_match = roof_new[
            (roof_new['Material'] == mat) &
            (roof_new['Code'] == row["Code"])
        ]
        if not roof_match.empty:
            thickness = roof_match.iloc[0]['Thickness(ft)']
            thicknesses_list.append(round(thickness, 2))
        else:
            print(f"WARNING: Thickness for material '{mat}' not found in Roof_New for construction '{row['Code']}'")
            thicknesses_list.append(0.0)

    # LAYER block
    mat_string = ', '.join([f'"{m}"' for m in material_names])
    thick_string = ', '.join(map(str, thicknesses_list))

    outputLayer += f'\n"{row["Code"]}_Lyr"  = LAYERS\n'
    outputLayer += f'  MATERIAL       = ({mat_string})\n'
    outputLayer += f'  THICKNESS      = ({thick_string})\n'
    outputLayer += f'  ..\n'

    # CONSTRUCTION block
    outputCons += f'"{row["Code"]}"  = CONSTRUCTION\n'
    outputCons += f'  TYPE           = LAYERS\n'
    outputCons += f'  LAYERS         = "{row["Code"]}_Lyr"\n'
    outputCons += f'  ..\n'

    # Insert updated blocks
    updated_inp_content = (
        inp_content[:end_index1] + outputMat +
        inp_content[end_index1:end_index2] + outputLayer + outputCons +
        inp_content[end_index2:]
    )

    # Replace only EXTERIOR-WALLs where LOCATION = TOP
    def replace_exterior_wall_construction(match):
        block = match.group(0)
        if 'LOCATION' in block and re.search(r'\bLOCATION\s*=\s*TOP\b', block):
            block = re.sub(r'(CONSTRUCTION\s*=\s*)"(.*?)"', rf'\1"{row["Code"]}"', block)
        return block

    updated_inp_content = re.sub(
        r'(".*?")\s*=\s*EXTERIOR-WALL\s*(.*?\.\.)',
        replace_exterior_wall_construction,
        updated_inp_content,
        flags=re.DOTALL
    )

    return updated_inp_content



# import pandas as pd
# import os
# import re

# def update_Material_Layers_Construction(inp_content, row_num):
#     # Load both sheets
#     matData = pd.read_excel('database/AllData.xlsx', sheet_name="Roof")
#     material_db = pd.read_excel('database/AllData.xlsx', sheet_name="Material_DB_IP")
   
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

#     matData = pd.read_excel('database/AllData.xlsx', sheet_name="Roof")

#     if row_num > len(matData):  # Use ">" instead of ">="
#         print(f"DEBUG: row_num={row_num}, matData length={len(matData)}")
#         raise IndexError(f"Row index {row_num} is out of bounds for matData with {len(matData)} rows.")

#     row = matData.iloc[row_num - 1]  # Convert from 1-based to 0-based index

#     outputMat = ""
#     outputLayer = ""
#     outputCons = ""

#     # Extracting relevant information for the current row
#     material_names = [row[f'Mat_{i}'] for i in range(1, 5)]
#     material_conduct = [row[f'Mat_{i}_conductivity'] for i in range(1, 5)]
#     material_density = [row[f'Mat_{i}_Density'] for i in range(1, 5)]
#     material_spc_heat = [row[f'Mat_{i}_Sp_Heat'] for i in range(1, 5)]
#     thicknesses = row['DefaultThick(ft)']

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
#     outputLayer += f'\n"{row["Code"]}_Lyr"  = LAYERS\n'
#     outputLayer += f'  MATERIAL       = ({material_str})\n'
#     outputLayer += f'  THICKNESS      = ({", ".join(map(str, thicknesses))})\n'
#     outputLayer += f'  ..\n'

#     outputCons += f'"{row["Code"]}"  = CONSTRUCTION\n'
#     outputCons += f'  TYPE           = LAYERS\n'
#     outputCons += f'  LAYERS         = "{row["Code"]}_Lyr"\n'
#     outputCons += f'  ..\n'

#     updated_inp_content = (
#         inp_content[:end_index1] + outputMat + inp_content[end_index1:end_index2] + outputLayer + outputCons + inp_content[end_index2:]
#     )

#     # # Replacing construction values in EXTERIOR-WALL sections
#     # updated_inp_content = re.sub(
#     #     r'(".*?")\s*=\s*EXTERIOR-WALL\s*'
#     #     r'CONSTRUCTION\s*=\s*"[^"]+"', 
#     #     lambda match: f'{match.group(1)} = EXTERIOR-WALL\n   CONSTRUCTION     = "{row["Code"]}"',
#     #     updated_inp_content
#     # )

#     def replace_exterior_wall_construction(match):
#             block = match.group(0)
#             name = match.group(1)
#             if 'LOCATION' in block and re.search(r'\bLOCATION\s*=\s*TOP\b', block):
#                 # Only replace CONSTRUCTION line
#                 block = re.sub(r'(CONSTRUCTION\s*=\s*)"(.*?)"', rf'\1"{row["Code"]}"', block)
#             return block
            
#     updated_inp_content = re.sub(
#         r'(".*?")\s*=\s*EXTERIOR-WALL\s*(.*?\.\.)',
#         replace_exterior_wall_construction,
#         updated_inp_content,
#         flags=re.DOTALL
#     )

#     return updated_inp_content

    
    # for index, row in matData.iterrows():
    #     outputMat = ""
    #     outputLayer = ""
    #     outputCons = ""

    #     # Extracting relevant information for the current row
    #     material_names = [row[f'Mat_{i}'] for i in range(1, 4)]
    #     material_conduct = [row[f'Mat_{i}_conductivity'] for i in range(1, 4)]
    #     material_density = [row[f'Mat_{i}_Density'] for i in range(1, 4)]
    #     material_spc_heat = [row[f'Mat_{i}_Sp_Heat'] for i in range(1, 4)]
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
    #     outputLayer += f'\n"{row["Code"]}_Lyr"  = LAYERS\n'
    #     outputLayer += f'  MATERIAL       = ({material_str})\n'
    #     outputLayer += f'  THICKNESS      = ({", ".join(map(str, thicknesses))})\n'
    #     outputLayer += f'  ..\n'

    #     outputCons += f'"{row["Code"]}"  = CONSTRUCTION\n'
    #     outputCons += f'  TYPE           = LAYERS\n'
    #     outputCons += f'  LAYERS         = "{row["Code"]}_Lyr"\n'
    #     outputCons += f'  ..\n'

    #     updated_inp_content = (
    #         inp_content[:end_index1] + outputMat + inp_content[end_index1:end_index2] + outputLayer + outputCons + inp_content[end_index2:]
    #     )

    #     def replace_exterior_wall_construction(match):
    #         block = match.group(0)
    #         name = match.group(1)
    #         if 'LOCATION' in block and re.search(r'\bLOCATION\s*=\s*TOP\b', block):
    #             # Only replace CONSTRUCTION line
    #             block = re.sub(r'(CONSTRUCTION\s*=\s*)"(.*?)"', rf'\1"{row["Code"]}"', block)
    #         return block
            
    #     updated_inp_content = re.sub(
    #         r'(".*?")\s*=\s*EXTERIOR-WALL\s*(.*?\.\.)',
    #         replace_exterior_wall_construction,
    #         updated_inp_content,
    #         flags=re.DOTALL
    #     )

    #     file_name = os.path.basename(name)
    #     file_name_without_extension, extension = os.path.splitext(file_name)

    #     output_file_name = f"{matData['Code'][index]}-{file_name_without_extension}.inp"
    #     output_file_path = os.path.join(os.path.dirname(name), output_file_name)

    #     with open(output_file_path, 'w') as file:
    #         file.write(updated_inp_content)