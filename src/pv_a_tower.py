import glob as gb
import os
import warnings
import pandas as pd
import xlwings as xw  # Xlwings is a Python library that makes it easy to call Python from Excel

warnings.filterwarnings("ignore")

def get_PVA_Tower_report(name, path):
    try:
        with open(name) as f:
            flist = f.readlines()
            
            pva_count = [] 
            for num, line in enumerate(flist, 0):
                if '*** COOLING TOWERS ***' in line:
                    pva_count.append(num)
                if 'PS-A' in line:
                    numend = num
                    break
                    
            if not pva_count:
                # Handle the case where '*** PRIMARY EQUIPMENT ***' is not found
                # print("The '*** PRIMARY EQUIPMENT ***' section was not found in the file:", name)
                return pd.DataFrame()  # Return an empty DataFrame
            
            numstart = pva_count[0] 
            pva_rpt = flist[numstart:numend]
            
            pva_str = []
            for line in pva_rpt:
                if('.' in line and ':' not in line and 'HTR' not in line):
                    pva_str.append(line)
                    
            result = []  
            for line in pva_str:
                pva_list = []
                splitter = line.split()
                space_name = "  ".join(splitter[:-8])
                pva_list=splitter[-8:]
                pva_list.insert(0,space_name)
                result.append(pva_list)
            pva_tower_df = pd.DataFrame(result)
            
            pva_tower_df.columns = ['EQUIPMENT TYPE', 'ATTACHED', 'TO', 'CAPACITY(MBTU/HR)',
                                'FLOW(GAL/MIN)', 'NUMBER_OF_CELLS', 'FAN_POWER_PER_CELL(KW)',
                                'FAN_PWR_PER_CELL(KW)', 'AUXILIARY(KW)']
            pva_tower_df['ATTACHED_TO'] = pva_tower_df['ATTACHED'].astype(str) + '_' + pva_tower_df['TO']
            pva_tower_df = pva_tower_df.drop(['ATTACHED', 'TO'], axis=1)
            last_column = pva_tower_df.iloc[:, -1]
            # Drop the last column from the DataFrame
            pva_tower_df = pva_tower_df.iloc[:, :-1]

            # Insert the last column to the first position
            pva_tower_df.insert(1, 'ATTACHED_TO', last_column)
            pva_tower_df.index.name = name
            value_before_backslash = ''.join(reversed(name)).split("\\")[0]
            name1 = ''.join(reversed(value_before_backslash))
            name = name1.rsplit(".", 1)[0]
            pva_tower_df.insert(0, 'RUNNAME', name)
        
        return pva_tower_df
    except Exception as e:
        columns = ['RUNNAME', 'HEATING_CAPACITY(MBTU/HR)', 'COOLING_CAPACITY(MBTU/HR)', 'LOOP_FLOW(GAL/MIN)',
                   'TOTAL_HEAD(FT)', 'SUPPLY_UA PRODUCT(BTU/HR-F)', 'SUPPLY_LOSS_DT(F)',
                   'RETURN_UA PRODUCT(BTU/HR-F)', 'RETURN_LOSS_DT(F)', 'LOOP_VOLUME(GAL)', 'FLUID_HEAT(CAPACITY)(BTU/LB-F)']
        return pd.DataFrame()