import glob as gb
import os
import warnings
import pandas as pd
import xlwings as xw # Xlwings is a Python library that makes it easy to call Python from Excel
import numpy as np

warnings.filterwarnings("ignore")

def get_SVA_Syst_report(name, path):
    try:
        with open(name) as f:
            flist = f.readlines()

            sva_count = [] 
            for num, line in enumerate(flist, 0):
                if 'SV-A' in line:
                    sva_count.append(num)
                if 'SS-D' in line:
                    numend = num
            numstart = sva_count[0] 
            sva_rpt = flist[numstart:numend]
            stop_values = [
                'SZRH', 'PSZ', 'SZCI', 'VAVS', 'PIU', 'CBVAV', 'RHFS', 'EVAP-COOL',
                'MZS', 'DDS', 'PMZS', 'FC', 'IU', 'FPH', 'PTAC', 'HP', 'HVSYS',
                'PVAVS', 'PVVT', 'UHT', 'UVT', 'RESYS', 'RESVVT']
            


            sva_str = []
            for line in sva_rpt:
                if (('PTAC' in line and '.' in line and 'WEATHER' not in line) or
                    ('VAVS' in line and '.' in line) or ('PIU' in line and '.' in line) or 
                    ('FC' in line and 'zn' not in line and 'Zn' not in line and '.' in line) or 
                    ('UVT' in line and '.' in line) or ('UHT' in line and '.' in line) or
                    ('PSZ' in line and '.' in line) or ('PMZS' in line and '.' in line) or ('SZRH' in line and '.' in line) or 
                    ('SZCI' in line and '.' in line) or ('CBVAV' in line and '.' in line) or ('RHFS' in line and '.' in line) or
                    ('EVAP-COOL' in line and '.' in line) or ('MZS' in line and '.' in line) or ('DDS' in line and '.' in line) or
                    ('IU' in line and '.' in line) or ('FPH' in line and '.' in line) or ('HP' in line and '.' in line) or
                    ('HVSYS' in line and '.' in line) or ('PVAVS' in line and '.' in line) or ('PVVT' in line and '.' in line) or 
                    ('RESYS' in line and '.' in line) or ('RESVVT' in line and '.' in line) or ('SUM' in line and '.' in line)):
                    sva_str.append(line)
            
            result = []  
            for line in sva_str:
                sva_list = []
                splitter = line.split()
                space_name = " ".join(splitter[:-10])
                sva_list=splitter[-10:]
                sva_list.insert(0,space_name)
                result.append(sva_list)
            sva_df = pd.DataFrame(result)
            sva_df.columns = ['SYSTEM_TYPE', 'ALTITUDE_FACTOR', 'FLOOR_AREA(SQFT)',
                            'MAX_PEOPLE', 'OUTSIDE_AIR_RATIO', 'COOLING_CAPACITY(KBTU/HR)',
                            'SENSIBLE(SHR)', 'HEATING_CAPACITY(KBTU/HR)', 'COOLING_EIR(BTU/BTU)', 'HEATING_EIR(BTU/BTU)',
                            'HEAT_PUMP(SUPP_HEAT)']
            sva_df['COOLING_CAPACITY(KBTU/HR)'] = pd.to_numeric(sva_df['COOLING_CAPACITY(KBTU/HR)'])
            sva_df['HEATING_CAPACITY(KBTU/HR)'] = pd.to_numeric(sva_df['HEATING_CAPACITY(KBTU/HR)'])
            sva_df['COOLING_EIR(BTU/BTU)'] = pd.to_numeric(sva_df['COOLING_EIR(BTU/BTU)'])
            sva_df['HEATING_EIR(BTU/BTU)'] = pd.to_numeric(sva_df['HEATING_EIR(BTU/BTU)'])
            
            sva_str1 = []
            for line in sva_rpt:
                if ('SUPPLY' in line and '.' in line):
                    sva_str1.append(line)
            
            result1 = []  
            for line in sva_str1:
                sva_list1 = []
                splitter1 = line.split()
                space_name1 = " ".join(splitter1[:-11])
                sva_list1=splitter1[-11:]
                sva_list1.insert(0,space_name1)
                result1.append(sva_list1)
            sva_df1 = pd.DataFrame(result1)
            sva_df1.columns = ['FAN_TYPE', 'CAPACITY(CFM)', 'DIVERSITY FACTOR(FRAC)',
                            'POWER DEMAND(KW)', 'FAN DELTA-T(F)', 'STATIC PRESSURE(IN WATER)',
                            'TOTAL EFF(FRAC)', 'MECH EFF(FRAC)', 'FAN PLACEMENT', 'FAN CONTROL', 'MAX FAN RATIO(FRAC)', 'MIN FAN RATIO(FRAC)']
            sva_df1['CAPACITY(CFM)'] = pd.to_numeric(sva_df1['CAPACITY(CFM)'])
            sva_df1['POWER DEMAND(KW)'] = pd.to_numeric(sva_df1['POWER DEMAND(KW)'])
            
            sva_df.reset_index(drop=True, inplace=True)
            sva_df1.reset_index(drop=True, inplace=True)
            result = pd.concat([sva_df, sva_df1], axis=1)
            
            result.drop(result.columns[[1, 2, 3, 4, 10, 13, 15, 16, 17, 18, 19, 20, 21, 22]], axis=1, inplace=True)
            result.drop(result.columns[[2, 6]], axis=1, inplace=True)
            
            ################################################
            sum_cooling_capacity = result.groupby('SYSTEM_TYPE')['COOLING_CAPACITY(KBTU/HR)'].sum()
            sum_heating_capacity = result.groupby('SYSTEM_TYPE')['HEATING_CAPACITY(KBTU/HR)'].sum()
            sum_cooling_eir = result.groupby('SYSTEM_TYPE')['COOLING_EIR(BTU/BTU)'].sum()
            sum_heating_eir = result.groupby('SYSTEM_TYPE')['HEATING_EIR(BTU/BTU)'].sum()
            sum_capacity = result.groupby('SYSTEM_TYPE')['CAPACITY(CFM)'].sum()
            sum_fan_power = result.groupby('SYSTEM_TYPE')['POWER DEMAND(KW)'].sum()
            ################################################

            data = {
                'SYSTEM_TYPE': sum_cooling_capacity.index,
                'COOLING_CAPACITY(KBTU/HR)': sum_cooling_capacity.values,
                'HEATING_CAPACITY(KBTU/HR)': sum_heating_capacity.values,
                'COOLING_EIR(BTU/BTU)': sum_cooling_eir.values,
                'HEATING_EIR(BTU/BTU)': sum_heating_eir.values,
                'CAPACITY(CFM)': sum_capacity.values,
                'POWER_DEMAND(KW)': sum_fan_power.values
            }

            final_df = pd.DataFrame(data)
            # print(final_df)
        
        return final_df
    except Exception as e:
        columns = ['RUNNAME', 'HEATING_CAPACITY(MBTU/HR)', 'COOLING_CAPACITY(MBTU/HR)', 'LOOP_FLOW(GAL/MIN)',
                   'TOTAL_HEAD(FT)', 'SUPPLY_UA PRODUCT(BTU/HR-F)', 'SUPPLY_LOSS_DT(F)',
                   'RETURN_UA PRODUCT(BTU/HR-F)', 'RETURN_LOSS_DT(F)', 'LOOP_VOLUME(GAL)', 'FLUID_HEAT(CAPACITY)(BTU/LB-F)']
        return pd.DataFrame()
