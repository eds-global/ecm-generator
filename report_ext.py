import os
import glob as gb
import pandas as pd
import shutil
from src import lv_b, ls_c, lv_d, pv_a_loop, sv_a, beps, bepu, lvd_summary, sva_zone, locationInfo, masterFile, sva_sys_type, pv_a_pump, pv_a_heater, pv_a_equip, pv_a_tower, ps_e, inp_shgc
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings

warnings.filterwarnings("ignore")

def generate_normalized_dataset(data):
    columns_to_fill = [
        "N-Wall-Area(SQFT)", "S-Wall-Area(SQFT)", "E-Wall-Area(SQFT)", "W-Wall-Area(SQFT)",
        "NE-Wall-Area(SQFT)", "NW-Wall-Area(SQFT)", "SE-Wall-Area(SQFT)", "SW-Wall-Area(SQFT)",
        "N-Wall-U-Value(BTU/HR-SQFT-F)", "S-Wall-U-Value(BTU/HR-SQFT-F)", "E-Wall-U-Value(BTU/HR-SQFT-F)",
        "W-Wall-U-Value(BTU/HR-SQFT-F)", "NE-Wall-U-Value(BTU/HR-SQFT-F)", "NW-Wall-U-Value(BTU/HR-SQFT-F)",
        "SE-Wall-U-Value(BTU/HR-SQFT-F)", "SW-Wall-U-Value(BTU/HR-SQFT-F)", 
        "N-Window-Area(SQFT)", "S-Window-Area(SQFT)", "E-Window-Area(SQFT)",
        "W-Window-Area(SQFT)", "NE-Window-Area(SQFT)", "NW-Window-Area(SQFT)", "SE-Window-Area(SQFT)",
        "SW-Window-Area(SQFT)", "N-Window-U-Value(BTU/HR-SQFT-F)",
        "S-Window-U-Value(BTU/HR-SQFT-F)", "E-Window-U-Value(BTU/HR-SQFT-F)", "W-Window-U-Value(BTU/HR-SQFT-F)",
        "NE-Window-U-Value(BTU/HR-SQFT-F)", "NW-Window-U-Value(BTU/HR-SQFT-F)", "SE-Window-U-Value(BTU/HR-SQFT-F)",
        "SW-Window-U-Value(BTU/HR-SQFT-F)"
    ]
    data[columns_to_fill] = data[columns_to_fill].fillna(0)

    # Derived columns
    data['TA'] = data['Floor-Total-Above-Grade(SQFT)'] + data['Floor-Total-Below-Grade(SQFT)']
    data['Power-Lighting(W/SQFT)'] = data['Power Lighting Total(W)'] / data['TA']
    data['Equipment-Tot(W/SQFT)'] = data['Equipment-Total(W)'] / data['TA']
    data['Roof-Conduction(KW)'] = data['ROOF CONDUCTION(KW)'] / data['ROOF-AREA(SQFT)']
    data['Wall-Conduction(KW)'] = data['WALL CONDUCTION(KW)'] / data['ALL WALLS-Wall-AREA(SQFT)']
    data['Energy_Outcome(KWH/SQFT)'] = data['Energy_Outcome(KWH)'] / data['TA']
    data['Total-LOAD(KW/SQFT)'] = data['TOTAL-LOAD(KW)'] / data['TA']
    data['Total-LOAD/Conditioned-Area(KW/SQFT)'] = data['TOTAL-LOAD(KW)'] / data['Floor-Total-Conditioned-Grade(SQFT)']
    data['Area*U (BTU/HR-F)'] = data['ALL WALLS-Wall-AREA(SQFT)'] * data['ALL WALLS-Wall-U-Value(BTU/HR-SQFT-F)']

    # Drop intermediate columns
    drop_cols = [
        'TA', 'Power Lighting Total(W)', 'Equipment-Total(W)', 'ROOF CONDUCTION(KW)', 
        'WALL CONDUCTION(KW)', 'Energy_Outcome(KWH)', 'TOTAL-LOAD(KW)'
    ]
    data.drop(columns=drop_cols, inplace=True)

    # Final column list
    columns = [
        "Batch_ID", "RunningUser", "SimulationLocation", "FileName", "ProjectCode", "ProjectName", "ProjectTypology", 
        "Location", "Climate", "Latitude", "Longitude", "Floor-Total-Above-Grade(SQFT)", 
        "Floor-Total-Below-Grade(SQFT)", "Floor-Total-Conditioned-Grade(SQFT)", 
        "Floor-Total-UnConditioned-Grade(SQFT)", "Wall-Total-Above-Grade(SQFT)", "Power-Lighting(W/SQFT)", 
        "Equipment-Tot(W/SQFT)", "N-Wall-Area(SQFT)", "S-Wall-Area(SQFT)", "E-Wall-Area(SQFT)", "W-Wall-Area(SQFT)",
        "NE-Wall-Area(SQFT)", "NW-Wall-Area(SQFT)", "SE-Wall-Area(SQFT)", "SW-Wall-Area(SQFT)", 
        "ROOF-AREA(SQFT)", "ALL WALLS-Wall-AREA(SQFT)", "WALLS+ROOFS-AREA(SQFT)", "UNDERGRND-Wall-AREA(SQFT)",
        "BUILDING-Wall-AREA(SQFT)", "N-Wall-U-Value(BTU/HR-SQFT-F)", "S-Wall-U-Value(BTU/HR-SQFT-F)", 
        "E-Wall-U-Value(BTU/HR-SQFT-F)", "W-Wall-U-Value(BTU/HR-SQFT-F)", "NE-Wall-U-Value(BTU/HR-SQFT-F)", 
        "NW-Wall-U-Value(BTU/HR-SQFT-F)", "SE-Wall-U-Value(BTU/HR-SQFT-F)", "SW-Wall-U-Value(BTU/HR-SQFT-F)", 
        "ROOF-U-Value(BTU/HR-SQFT-F)", "ALL WALLS-Wall-U-Value(BTU/HR-SQFT-F)", "WALLS+ROOFS-U-Value(BTU/HR-SQFT-F)", 
        "UNDERGRND-Wall-U-Value(BTU/HR-SQFT-F)", "BUILDING-Wall-U-Value(BTU/HR-SQFT-F)", "N-Window-Area(SQFT)", 
        "S-Window-Area(SQFT)", "E-Window-Area(SQFT)", "W-Window-Area(SQFT)", "NE-Window-Area(SQFT)", 
        "NW-Window-Area(SQFT)", "SE-Window-Area(SQFT)", "SW-Window-Area(SQFT)", "ROOF-Window-Area(SQFT)", 
        "ALL WALLS-Window-AREA(SQFT)", "WALLS+ROOFS-Window-AREA(SQFT)", "UNDERGRND-Window-AREA(SQFT)", 
        "BUILDING-Window-AREA(SQFT)", "N-Window-U-Value(BTU/HR-SQFT-F)", "S-Window-U-Value(BTU/HR-SQFT-F)", 
        "E-Window-U-Value(BTU/HR-SQFT-F)", "W-Window-U-Value(BTU/HR-SQFT-F)", "NE-Window-U-Value(BTU/HR-SQFT-F)", 
        "NW-Window-U-Value(BTU/HR-SQFT-F)", "SE-Window-U-Value(BTU/HR-SQFT-F)", "SW-Window-U-Value(BTU/HR-SQFT-F)", 
        "ROOF-Window-U-Value(BTU/HR-SQFT-F)", "ALL WALLS-Window-U-Value(BTU/HR-SQFT-F)", 
        "WALLS+ROOFS-Window-U-Value(BTU/HR-SQFT-F)", "UNDERGRND-Window-U-Value(BTU/HR-SQFT-F)", 
        "BUILDING-Window-U-Value(BTU/HR-SQFT-F)", "WWR", "Area*U (BTU/HR-F)", "Wall-Conduction(KW)", 
        "Roof-Conduction(KW)", "WINDOW GLASS+FRM COND(KW)", "WINDOW GLASS SOLAR(KW)", "DOOR CONDUCTION(KW)", 
        "INTERNAL SURFACE COND(KW)", "UNDERGROUND SURF COND(KW)", "OCCUPANTS TO SPACE(KW)", 
        "LIGHT TO SPACE(KW)", "EQUIPMENT TO SPACE(KW)", "PROCESS TO SPACE(KW)", "INFILTRATION(KW)", 
        "Total-LOAD(KW/SQFT)", "Total-LOAD/Conditioned-Area(KW/SQFT)", "Energy_Outcome(KWH/SQFT)"
    ]

    # Filter desired columns
    new_data = data[columns]

    # Additional feature engineering
    new_data['Above-Grade/Below-Grade'] = new_data["Floor-Total-Above-Grade(SQFT)"] / new_data["Floor-Total-Below-Grade(SQFT)"]
    new_data['Conditioned-Area/UnConditioned-Area'] = new_data["Floor-Total-Conditioned-Grade(SQFT)"] / new_data["Floor-Total-UnConditioned-Grade(SQFT)"]
    new_data['Total-Above-Grade-Ext-Wall-Area/Total-AG-FloorArea'] = new_data["Wall-Total-Above-Grade(SQFT)"] / new_data["Floor-Total-Above-Grade(SQFT)"]
    new_data['Roof-Window-Area/Roof-Area'] = new_data["ROOF-Window-Area(SQFT)"] / new_data["ROOF-AREA(SQFT)"]
    new_data['Roof-Area/Total-AG-Floor-Area'] = new_data["ROOF-AREA(SQFT)"] / new_data["Floor-Total-Above-Grade(SQFT)"]

    # Columns to drop
    columns_to_drop = [
        "SimulationLocation", "ProjectCode", "ProjectName", "Location", "Latitude", "Longitude", 
        "Floor-Total-Above-Grade(SQFT)", "Floor-Total-Below-Grade(SQFT)", "Floor-Total-Conditioned-Grade(SQFT)", 
        "Floor-Total-UnConditioned-Grade(SQFT)", "Wall-Total-Above-Grade(SQFT)", "N-Wall-Area(SQFT)", 
        "S-Wall-Area(SQFT)", "E-Wall-Area(SQFT)", "W-Wall-Area(SQFT)", "NE-Wall-Area(SQFT)", "NW-Wall-Area(SQFT)", 
        "SE-Wall-Area(SQFT)", "SW-Wall-Area(SQFT)", "ROOF-AREA(SQFT)", "ALL WALLS-Wall-AREA(SQFT)", 
        "WALLS+ROOFS-AREA(SQFT)", "UNDERGRND-Wall-AREA(SQFT)", "BUILDING-Wall-AREA(SQFT)", 
        "N-Wall-U-Value(BTU/HR-SQFT-F)", "S-Wall-U-Value(BTU/HR-SQFT-F)", "E-Wall-U-Value(BTU/HR-SQFT-F)", 
        "W-Wall-U-Value(BTU/HR-SQFT-F)", "NE-Wall-U-Value(BTU/HR-SQFT-F)", "NW-Wall-U-Value(BTU/HR-SQFT-F)", 
        "SE-Wall-U-Value(BTU/HR-SQFT-F)", "SW-Wall-U-Value(BTU/HR-SQFT-F)", 
        "WALLS+ROOFS-U-Value(BTU/HR-SQFT-F)", "BUILDING-Wall-U-Value(BTU/HR-SQFT-F)", 
        "N-Window-Area(SQFT)", "S-Window-Area(SQFT)", "E-Window-Area(SQFT)", "W-Window-Area(SQFT)", 
        "NE-Window-Area(SQFT)", "NW-Window-Area(SQFT)", "SE-Window-Area(SQFT)", "SW-Window-Area(SQFT)", 
        "ROOF-Window-Area(SQFT)", "ALL WALLS-Window-AREA(SQFT)", "WALLS+ROOFS-Window-AREA(SQFT)", 
        "UNDERGRND-Window-AREA(SQFT)", "BUILDING-Window-AREA(SQFT)", 
        "N-Window-U-Value(BTU/HR-SQFT-F)", "S-Window-U-Value(BTU/HR-SQFT-F)", "E-Window-U-Value(BTU/HR-SQFT-F)", 
        "W-Window-U-Value(BTU/HR-SQFT-F)", "NE-Window-U-Value(BTU/HR-SQFT-F)", "NW-Window-U-Value(BTU/HR-SQFT-F)", 
        "SE-Window-U-Value(BTU/HR-SQFT-F)", "SW-Window-U-Value(BTU/HR-SQFT-F)", 
        "WALLS+ROOFS-Window-U-Value(BTU/HR-SQFT-F)", "UNDERGRND-Window-U-Value(BTU/HR-SQFT-F)", 
        "BUILDING-Window-U-Value(BTU/HR-SQFT-F)", "Area*U (BTU/HR-F)", "Wall-Conduction(KW)", 
        "Roof-Conduction(KW)", "WINDOW GLASS+FRM COND(KW)", "WINDOW GLASS SOLAR(KW)", 
        "DOOR CONDUCTION(KW)", "INTERNAL SURFACE COND(KW)", "UNDERGROUND SURF COND(KW)", 
        "OCCUPANTS TO SPACE(KW)", "LIGHT TO SPACE(KW)", "EQUIPMENT TO SPACE(KW)", 
        "PROCESS TO SPACE(KW)", "INFILTRATION(KW)"
    ]

    new_data = new_data.drop(columns=columns_to_drop, errors='ignore')

    # Reorder
    desired_order = [
        "Batch_ID", "RunningUser", "FileName", "ProjectTypology", "Climate", "Above-Grade/Below-Grade", 
        "Conditioned-Area/UnConditioned-Area", "Roof-Area/Total-AG-Floor-Area",
        "Roof-Window-Area/Roof-Area", "Total-Above-Grade-Ext-Wall-Area/Total-AG-FloorArea",
        "Power-Lighting(W/SQFT)", "Equipment-Tot(W/SQFT)", "ROOF-U-Value(BTU/HR-SQFT-F)",
        "ALL WALLS-Wall-U-Value(BTU/HR-SQFT-F)", "UNDERGRND-Wall-U-Value(BTU/HR-SQFT-F)", 
        "ROOF-Window-U-Value(BTU/HR-SQFT-F)", "ALL WALLS-Window-U-Value(BTU/HR-SQFT-F)",
        "WWR", "Total-LOAD(KW/SQFT)", "Total-LOAD/Conditioned-Area(KW/SQFT)", "Energy_Outcome(KWH/SQFT)"
    ]
    new_data = new_data[desired_order]

    # Clean infinities
    new_data["Above-Grade/Below-Grade"].replace([np.inf, -np.inf], 0, inplace=True)

    # Export to Excel
    # new_data.to_excel(output_excel, index=False, engine='openpyxl')
    # print("Normalized dataset generated and saved to:", output_excel)
    return new_data

# function to get report in csv and save in specific folders
def get_report_and_save(report_function, name, file_suffix, path):
    try:
        #print("calling report function: ", report_function, name, path)
        report = report_function(name, path)
        name = os.path.splitext(name)[0]
        # get file path name as .csv
        file_path = os.path.join(path, f'{name}_{file_suffix}.csv')
        # if that file already exist, replace with other file.
        if os.path.isfile(file_path):
            os.remove(file_path)
        # writing csv file with headers and no index column
        print("Generating -",file_suffix, "....")
        with open(file_path, 'w', newline='') as f:
            report.to_csv(f, header=True, index=False, mode='wt')
    except:
        print(f"Skipping {file_suffix}...")
     


#get success files from log file
def get_files_for_data_extraction(output_path, log_file_path, new_batch_id, location_id, user_nm):
    df = pd.read_excel(log_file_path)
     # Filter to only 'Fail' status
    success_files = df.loc[df['Status'] == 'Success', 'File Name'].tolist()
    #print(success_files)

    #iterate through success file list
    i = 1
    for file in success_files:
        print (f"\nExtracting file{i}: {file} *********")
        i = i+1
        #get file path from file name
        sim_file_path = os.path.join(output_path, file)
        inp_file = file.replace('.sim', '.inp')
        inp_file_path = os.path.join(output_path,inp_file)

        name = os.path.basename(file)
        
        #print(name)

        #extracting data from sim file
        try:
            get_report_and_save(ls_c.get_LSC_report, sim_file_path, 'lsc', output_path)
            get_report_and_save(lv_d.get_LVD_report, sim_file_path, 'lvd', output_path)
            get_report_and_save(lvd_summary.get_LVD_Summary_report, sim_file_path, 'lvd_Summary', output_path)
            get_report_and_save(pv_a_loop.get_PVA_report, sim_file_path, 'pva_loop', output_path)
            get_report_and_save(sv_a.get_SVA_report, sim_file_path, 'sva', output_path)
            get_report_and_save(sva_zone.get_SVA_Zone_report, sim_file_path, 'sva_Zone', output_path)
            get_report_and_save(beps.get_BEPS_report, sim_file_path, 'beps', output_path)
            get_report_and_save(bepu.get_BEPU_report, sim_file_path, 'bepu', output_path)
            get_report_and_save(sva_sys_type.get_SVA_Syst_report, sim_file_path, 'sva_syst', output_path)
            get_report_and_save(pv_a_pump.get_PVA_Pump_report, sim_file_path, 'pva_pump', output_path)
            get_report_and_save(pv_a_heater.get_PVA_Heater_report, sim_file_path, 'pva_heater', output_path)
            get_report_and_save(pv_a_equip.get_PVA_Equip_report, sim_file_path, 'pva_equip', output_path)
            get_report_and_save(pv_a_tower.get_PVA_Tower_report, sim_file_path, 'pva_tower', output_path)
            get_report_and_save(ps_e.get_PSE_report, sim_file_path, 'pse', output_path)
            get_report_and_save(lv_b.get_LVB_report, sim_file_path, 'lvb', output_path)
        except Exception as e:
            print(f"Skipping!!")
        
        #getting data from inp file
        try:
            name = os.path.basename(inp_file)
            
            get_report_and_save(inp_shgc.get_SHGC_report, inp_file_path, 'shgc', output_path)
        except Exception as e:
            print(f"Skipping INP!!")
        
    print("extraction complete")
        

    
    combined_data = masterFile.get_all_calculated_values(locationInfo.get_locInfo(output_path), output_path)
    combined_data.insert(loc=0, column="Batch_ID", value=new_batch_id)
    combined_data.insert(loc=1, column="RunningUser", value=user_nm)
    combined_data.insert(loc=2, column="SimulationLocation", value=location_id)
    df = combined_data.copy()
    
    Normalized_Data = generate_normalized_dataset(df)

    if combined_data is None:
         print("combined_data is None, saving empty CSV.")
    else:
        combined_folder_path = os.path.dirname(output_path)
        combined_file_path = os.path.join(combined_folder_path, 'combined_data.csv')
        combined_file_path_normalized = os.path.join(combined_folder_path, 'Normalized_Data.csv')
        if os.path.exists(combined_file_path):
            combined_data.to_csv(os.path.join(combined_file_path), mode='a', header=False, index=False)
            # Normalized_Data.to_csv(os.path.join(combined_file_path_normalized), mode='a', header=False, index=False)
        else:
            combined_data.to_csv(os.path.join(combined_file_path), index=False)
            # Normalized_Data.to_csv(os.path.join(combined_file_path_normalized), index=False)
        print(f"✅ CSVs generated at:\n- {combined_file_path}")