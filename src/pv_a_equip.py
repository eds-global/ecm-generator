import glob as gb
import os
import warnings
import pandas as pd
import xlwings as xw  # Xlwings is a Python library that makes it easy to call Python from Excel

warnings.filterwarnings("ignore")

def get_PVA_Equip_report(name, path):
    try:
        with open(name) as f:
            flist = f.readlines()
            
            pva_count = [] 
            for num, line in enumerate(flist, 0):
                if '*** PRIMARY EQUIPMENT ***' in line:
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
                if('.' in line and ':' not in line and 'TWR' not in line and 'HEATER' not in line and 'Condenser' not in line and
                'HTR' not in line):
                    pva_str.append(line)
                    
            result = []  
            for line in pva_str:
                pva_list = []
                splitter = line.split()
                space_name = "  ".join(splitter[:-7])
                pva_list=splitter[-7:]
                pva_list.insert(0,space_name)
                result.append(pva_list)
            pva_equip_df = pd.DataFrame(result)
            
            pva_equip_df.columns = ['EQUIPMENT TYPE', 'ATTACHED', 'TO', 'RATED CAPACITY(MBTU/HR)',
                                'FLOW(GAL/MIN)', 'RATED EIR(FRAC)', 'RATED HIR(FRAC)',
                                'AUXILIARY(KW)']
            pva_equip_df['ATTACHED_TO'] = pva_equip_df['ATTACHED'].astype(str) + '_' + pva_equip_df['TO']
            pva_equip_df = pva_equip_df.drop(['ATTACHED', 'TO'], axis=1)
            last_column = pva_equip_df.iloc[:, -1]
            # Drop the last column from the DataFrame
            pva_equip_df = pva_equip_df.iloc[:, :-1]

            # Insert the last column to the first position
            pva_equip_df.insert(1, 'ATTACHED_TO', last_column)
            pva_equip_df.index.name = name
            value_before_backslash = ''.join(reversed(name)).split("\\")[0]
            name1 = ''.join(reversed(value_before_backslash))
            name = name1.rsplit(".", 1)[0]
            pva_equip_df.insert(0, 'RUNNAME', name)
        
        return pva_equip_df
    except Exception as e:
        columns = ['RUNNAME', 'HEATING_CAPACITY(MBTU/HR)', 'COOLING_CAPACITY(MBTU/HR)', 'LOOP_FLOW(GAL/MIN)',
                   'TOTAL_HEAD(FT)', 'SUPPLY_UA PRODUCT(BTU/HR-F)', 'SUPPLY_LOSS_DT(F)',
                   'RETURN_UA PRODUCT(BTU/HR-F)', 'RETURN_LOSS_DT(F)', 'LOOP_VOLUME(GAL)', 'FLUID_HEAT(CAPACITY)(BTU/LB-F)']
        return pd.DataFrame()