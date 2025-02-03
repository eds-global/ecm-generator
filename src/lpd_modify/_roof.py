import pandas as pd
import os
import streamlit as st

def update_Material_Layers_Construction_Roof(name, roof):
    if roof is None:
        return name

    with open(name, 'r') as file:
        inp_content = file.read()

    start_marker = "Materials / Layers / Constructions"
    end_marker1 = "= LAYERS"
    end_marker2 = "= CONSTRUCTION"
    
    start_index = inp_content.find(start_marker)
    end_index1 = inp_content.find(end_marker1, start_index)
    end_index2 = inp_content.find(end_marker2, start_index)
    
    if start_index == -1 or end_index1 == -1 or end_index2 == -1:
        raise ValueError("Markers not found in the input file.")
    
    end_index1 = inp_content.rfind('\n', start_index, end_index1)
    end_index2 = inp_content.rfind('\n', start_index, end_index2)

    matData = pd.read_excel('database/ML_ScaleUp_v02.xlsx', sheet_name="Roof-Options")
    selected_row = matData[matData['Roof-Name'] == roof]

    if selected_row.empty:
        print(f"No matching data found for wall: {roof}")
        return inp_content

    row = selected_row.iloc[0]

    materials = []
    for i in range(1, 4):
        material_name = row[f'Mat_{i}']
        if material_name not in materials:
            materials.append(material_name)
    
    material_conduct = [row[f'Mat_{i}_conductivity'] for i in range(1, 4)]
    material_density = [row[f'Mat_{i}_Density'] for i in range(1, 4)]
    material_spc_heat = [row[f'Mat_{i}_Sp_Heat'] for i in range(1, 4)]
    thicknesses = [round(row[f'Mat_{i}_thickness'], 2) for i in range(1, 4)]
    
    outputMat = ""
    for i, material in enumerate(materials):
        outputMat += (
            f'"{material}"  = MATERIAL\n'
            f'  TYPE           = PROPERTIES\n'
            f'  THICKNESS      = {thicknesses[i]}\n'
            f'  CONDUCTIVITY   = {material_conduct[i]:.2f}\n'
            f'  DENSITY        = {material_density[i]:.2f}\n'
            f'  SPECIFIC-HEAT  = {material_spc_heat[i]:.2f}\n'
            f'  ..\n'
        )
    
    material_str = ', '.join([f'"{m}"' for m in materials])
    outputLayer = (
        f'"{row["Code"]}_Lyr"  = LAYERS\n'
        f'  MATERIAL       = ({material_str})\n'
        f'  THICKNESS      = ({", ".join(map(str, thicknesses[:len(materials)]))})\n'
        f'  ..\n'
    )

    outputCons = (
        f'"{row["Code"]}"  = CONSTRUCTION\n'
        f'  TYPE           = LAYERS\n'
        f'  LAYERS         = "{row["Code"]}_Lyr"\n'
        f'  ..\n'
    )

    updated_inp_content = (
        inp_content[:end_index1] + outputMat + inp_content[end_index1:end_index2] + outputLayer + outputCons + inp_content[end_index2:]
    )

    with open(name, 'w') as file:
        file.write(updated_inp_content)

    return name