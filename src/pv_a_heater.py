import glob as gb
import os
import warnings
import pandas as pd
import xlwings as xw  # Xlwings is a Python library that makes it easy to call Python from Excel

warnings.filterwarnings("ignore")

def get_PVA_Heater_report(name, path):
    try:
        with open(name) as f:
            flist = f.readlines()
            
            pva_count = [] 
            for num, line in enumerate(flist, 0):
                if '*** DW-HEATERS ***' in line:
                    pva_count.append(num)
                if 'PS-A' in line:
                    numend = num
                    break

            if not pva_count:
                # Handle the case where '*** PUMPS ***' is not found
                # print("The '*** PUMPS ***' section was not found in the file:", name)
                return pd.DataFrame()  # Return an empty DataFrame
            
            numstart = pva_count[0] 
            pva_rpt = flist[numstart:numend]
            
            pva_str = []
            for line in pva_rpt:
                if('.' in line and ':' not in line):
                    pva_str.append(line)
                    
            result = []  
            for line in pva_str:
                pva_list = []
                splitter = line.split()
                space_name = " ".join(splitter[:-7])
                pva_list=splitter[-7:]
                pva_list.insert(0,space_name)
                result.append(pva_list)
            pva_heater_df = pd.DataFrame(result)
            
            pva_heater_df.columns = ['EQUIPMENT TYPE_ATTACHED TO', 'CAPACITY(MBTU/HR)',
                                'FLOW(GAL/MIN)', 'EIR(FRAC)', 'HIR(FRAC)',
                                'AUXILIARY(KW)', 'TANK(GAL)', 'TANK UA(BTU/HR-F)']
            pva_heater_df.index.name = name
            value_before_backslash = ''.join(reversed(name)).split("\\")[0]
            name1 = ''.join(reversed(value_before_backslash))
            name = name1.rsplit(".", 1)[0]
            pva_heater_df.insert(0, 'RUNNAME', name)
                    
        return pva_heater_df
    except Exception as e:
        columns = ['RUNNAME', 'HEATING_CAPACITY(MBTU/HR)', 'COOLING_CAPACITY(MBTU/HR)', 'LOOP_FLOW(GAL/MIN)',
                   'TOTAL_HEAD(FT)', 'SUPPLY_UA PRODUCT(BTU/HR-F)', 'SUPPLY_LOSS_DT(F)',
                   'RETURN_UA PRODUCT(BTU/HR-F)', 'RETURN_LOSS_DT(F)', 'LOOP_VOLUME(GAL)', 'FLUID_HEAT(CAPACITY)(BTU/LB-F)']
        return pd.DataFrame()