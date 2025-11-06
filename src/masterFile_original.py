import glob as gb
import os
import warnings
import pandas as pd
import xlwings as xw

warnings.filterwarnings("ignore")

def get_AFR_FanPower_Cool_Heat_Capacities_EIRs(name, x):
    calc_database = pd.read_csv("database/eQuest Automation System Mapping.csv")

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

        sva_str = []
        for line in sva_rpt:
            if (('SZRH' in line and '.' in line) or ('PSZ' in line and '.' in line) or 
                ('SZCI' in line and '.' in line) or ('VAVS' in line and '.' in line) or 
                ('PIU' in line and '.' in line) or ('CBVAV' in line and '.' in line) or 
                ('RHFS' in line and '.' in line) or ('EVAP-COOL' in line and '.' in line) or 
                ('MZS' in line and '.' in line) or ('DDS' in line and '.' in line) or 
                ('PMZS' in line and '.' in line) or 
                ('FC' in line and 'zn' not in line and 'Zn' not in line and '.' in line) or 
                ('IU' in line and '.' in line) or ('FPH' in line and '.' in line) or 
                ('PTAC' in line and '.' in line and 'WEATHER' not in line) or 
                ('HP' in line and '.' in line) or ('HVSYS' in line and '.' in line) or 
                ('PVAVS' in line and '.' in line) or ('PVVT' in line and '.' in line) or 
                ('UHT' in line and '.' in line) or ('UVT' in line and '.' in line) or 
                ('RESYS' in line and '.' in line) or ('RESVVT' in line and '.' in line) and 'FILE-' not in line):
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
        sva_df.columns = ['SYSTEM_TYPE', 'ALTITUDE_FACTOR', 'SYS_FLOOR_AREA', 'SYS_MAX_PEOPLE', 
                        'SYS_OUTSIDE_AIR_RATIO', 'SYS_COOLING_CAPACITY', 'SYS_SENSIBLE', 'SYS_HEATING_CAPACITY',
                        'SYS_COOLING_EIR', 'SYS_HEATING_EIR', 'HEAT_PUMP_SUPP_HEAT']
                
        sva_str1 = []
        for line in sva_rpt:
            if (('SUPPLY' in line and '.' in line)):
                sva_str1.append(line)
        
        result1 = []
        for line in sva_str1:
            sva_list1 = []
            splitter1 = line.split()
            
            if 'BY' in line:
                space_name1 = " ".join(splitter1[:-12])
                sva_list1 = splitter1[-12:]
                sva_list1.insert(0, space_name1)
                
                if len(sva_list1) == 13:  # Ensure it matches the expected length
                    result1.append(sva_list1)
            else:
                space_name1 = " ".join(splitter1[:-11])
                sva_list1 = splitter1[-11:]
                sva_list1.insert(0, space_name1)
                
                if len(sva_list1) == 12:  # Ensure it matches the expected length
                    result1.append(sva_list1)

        # Create DataFrame and assign columns based on length
        sva_df1 = pd.DataFrame(result1)

        if sva_df1.shape[1] == 13:
            sva_df1.columns = ['FAN_TYPE', 'CAPACITY', 'DIVERSITY_FACTOR',
                            'POWER_DEMAND', 'FAN-DELTA-T', 'STATIC_PRESSURE',
                            'TOTAL_EFF', 'MECH_EFF', 'FAN_PLACEMENT', 
                            'FAN_CONTROL', '1', 'MAX_FAN_RATIO', 'MIN_FAN_RATIO']
            sva_df1['FAN_CONTROL'] = sva_df1['FAN_CONTROL'] + sva_df1['1'].astype(str)
            sva_df1.drop(columns=['1'], inplace=True)
        elif sva_df1.shape[1] == 12:
            sva_df1.columns = ['FAN_TYPE', 'CAPACITY', 'DIVERSITY_FACTOR',
                            'POWER_DEMAND', 'FAN-DELTA-T', 'STATIC_PRESSURE',
                            'TOTAL_EFF', 'MECH_EFF', 'FAN_PLACEMENT', 
                            'FAN_CONTROL', 'MAX_FAN_RATIO', 'MIN_FAN_RATIO']
        else:
            raise ValueError(f"Unexpected number of columns: {sva_df1.shape[1]}")
                # sva_df1['FAN_CONTROL'] = sva_df1['FAN_CONTROL'] + sva_df1['1'].astype(str)
                # sva_df1.drop(columns=['1'], inplace=True)
        
        # Reset the indexes of all the DataFrames
        sva_df.reset_index(drop=True, inplace=True)
        sva_df1.reset_index(drop=True, inplace=True)
        
        # Concatenate the DataFrames along the columns axis
        sva_result = pd.concat([sva_df, sva_df1], axis=1)
        sva_result.reset_index(drop=True, inplace=True)
        sva_result.drop(sva_result.columns[[1, 2, 3, 4, 6, 10, 13, 15, 16, 17, 18, 19, -1, -2, -3]], axis=1, inplace=True)
        sva_result.drop(sva_result.columns[[5]], axis=1, inplace=True) 
                
        sva_str2 = []
        for line in sva_rpt:
            if ('SUPPLY' not in line and '.' in line and ('SZRH' in line and '.' in line) or ('PSZ' in line and '.' in line) or 
                ('SZCI' in line and '.' in line) or ('VAVS' in line and '.' in line) or 
                ('PIU' in line and '.' in line) or ('CBVAV' in line and '.' in line) or 
                ('RHFS' in line and '.' in line) or ('EVAP-COOL' in line and '.' in line) or 
                ('MZS' in line and '.' in line) or ('DDS' in line and '.' in line) or 
                ('PMZS' in line and '.' in line) or ('SUM' in line and '.' in line) or
                ('FC' in line and '.' in line) or 
                ('IU' in line and '.' in line) or ('FPH' in line and '.' in line) or 
                ('PTAC' in line and '.' in line and 'WEATHER' not in line) or 
                ('HP' in line and '.' in line) or ('HVSYS' in line and '.' in line) or 
                ('PVAVS' in line and '.' in line) or ('PVVT' in line and '.' in line) or 
                ('UHT' in line and '.' in line) or ('UVT' in line and '.' in line) or 
                ('RESYS' in line and '.' in line) or ('RESVVT' in line and '.' in line) or ('Zn' in line and '.' in line) or
            ('zn' in line and '.' in line) or ('zone' in line and '.' in line) or('Zone' in line and '.' in line)):
                sva_str2.append(line)
                
        result2 = []  
        for line in sva_str2:
            sva_list2 = []
            splitter2 = line.split()
            space_name2 = " ".join(splitter2[:-11])
            sva_list2=splitter2[-11:]
            sva_list2.insert(0,space_name2)
            result2.append(sva_list2)

        sva_df2 = pd.DataFrame(result2)
        
        sva_df2.columns = ['zn','A', 'B',
                        'C', 'D', 'E',
                        'F', 'G', 'H', 'I', 'J', 'K']
        # Iterate over rows
        drop_rows = False
        indices_to_drop = []
        stop_values = [
            'SZRH', 'PSZ', 'SZCI', 'VAVS', 'PIU', 'CBVAV', 'RHFS', 'EVAP-COOL',
            'MZS', 'DDS', 'PMZS', 'FC', 'IU', 'FPH', 'PTAC', 'HP', 'HVSYS',
            'PVAVS', 'PVVT', 'UHT', 'UVT', 'RESYS', 'RESVVT']
        
        for index, row in sva_df2.iterrows():
            if row['A'] == 'SUM' or row['A'] == '':
                drop_rows = True
            elif row['A'] in stop_values:
                drop_rows = False
            if drop_rows:
                indices_to_drop.append(index)
                
        # Drop rows
        sva_df2 = sva_df2.drop(indices_to_drop)
        sva_df2 = sva_df2[sva_df2['B'] != 'zn']
        sva_df2.drop(sva_df2.columns[[0]], axis=1, inplace=True)
        sva_df2.drop(sva_df2.columns[[1, 3, 4, 6]], axis=1, inplace=True) 
        # Iterate over each row
        for index, row in sva_df2.iterrows():
            # Check if the first column of the row contains any alphabet characters
            if any(c.isalpha() for c in row['A']):
                # If the condition is met, make all columns in that row the same as the value in the first column
                string_value = row['A']
                sva_df2.loc[index] = [string_value] * len(row)

        a = sva_df2['A'].to_list()
        A = []
        b = sva_df2['C'].to_list()
        B = []
        c = sva_df2['F'].to_list()
        C = []
        d = sva_df2['H'].to_list()
        D = []
        e = sva_df2['I'].to_list()
        E = []
        f = sva_df2['J'].to_list()
        F = []
        g = sva_df2['K'].to_list()
        G = []
                
        sum1 = 0
        for (i) in range(1, len(a)):
            if '.' in a[i]:
                sum1 += float(a[i])
            else:
                A.append(sum1)
                sum1 = 0
        if sum1 != 0:
            A.append(sum1)
            
        sum2 = 0
        for (i) in range(1, len(b)):
            if '.' in b[i]:
                sum2 += float(b[i])
            else:
                B.append(sum2)
                sum2 = 0
        if sum2 != 0:
            B.append(sum2)

        sum3 = 0
        for (i) in range(1, len(c)):
            if '.' in c[i]:
                sum3 += float(c[i])
            else:
                C.append(sum3)
                sum3 = 0
        if sum3 != 0:
            C.append(sum3)
            
        sum4 = 0
        for (i) in range(1, len(d)):
            if '.' in d[i]:
                sum4 += float(d[i])
            else:
                D.append(sum4)
                sum4 = 0
        if sum4 != 0:
            D.append(sum4)
            
        sum5 = 0
        for (i) in range(1, len(e)):
            if '.' in e[i]:
                sum5 += float(e[i])
            else:
                E.append(sum5)
                sum5 = 0
        if sum5 != 0:
            E.append(sum5)
            
        sum6 = 0
        for (i) in range(1, len(f)):
            if '.' in f[i]:
                sum6 += float(f[i])
            else:
                F.append(sum6)
                sum6 = 0
        if sum6 != 0:
            F.append(sum6)
                
        sum7 = 0
        for (i) in range(1, len(g)):
            if '.' in g[i]:
                sum7 += float(g[i])
            else:
                G.append(sum7)
                sum7 = 0
        if sum7 != 0:
            G.append(sum7)
            
            # Determine the length of the shorter list
        max_length = min(len(A), len(B), len(C), len(D), len(E), len(F), len(G))

        # Creating a DataFrame
        data = {
            'SUPPLY_FLOW': A[:max_length],
            'FAN': B[:max_length],
            'ZONE_COOLING_CAPACITY': C[:max_length],
            'ZONE_EXTRACTION': D[:max_length],
            'ZONE_HEATING_CAPACITY': E[:max_length],
            'ZONE_ADDITION_RATE': F[:max_length],
            'ZONE_MULT': G[:max_length]
        }

        # Create the DataFrame using the determined length
        sva_df3 = pd.DataFrame(data)
        
        # Reset the indexes of all the DataFrames
        sva_df3.reset_index(drop=True, inplace=True)
        sva_result.reset_index(drop=True, inplace=True)
        
        # Concatenate the DataFrames along the columns axis
        sva_result = pd.concat([sva_result, sva_df3], axis=1)
        
#################################### CALCULATIONS STARTS FROM HERE ##################################
        # List of system types to drop for calaculating cooling_eir
        system_types_to_drop1 = ['SZRH', 'SZCI', 'VAVS', 'CBVAV', 'RHFS', 'MZS', 'DDS', 'FC', 'IU', 'EVAP-COOL', 'FPH', 'HVSYS', 'UHT', 'UVT', 'SUM']
        # List of system types to drop for calaculating cooling_capacity
        system_types_to_drop2 = ['EVAP-COOL', 'FPH', 'HVSYS', 'UHT', 'UVT', 'SUM']
        # List of system types to drop for calaculating cooling_capacity
        system_types_to_drop3 = ['FPH', 'SUM']
        # List of system types to drop for calaculating Heating_capacity
        system_types_to_drop4 = ['SUM']
        # Filter out rows based on system types to drop
        sva_result1 = sva_result[~sva_result['SYSTEM_TYPE'].isin(system_types_to_drop1)]
        # Filter out rows based on system types to drop
        sva_result2 = sva_result[~sva_result['SYSTEM_TYPE'].isin(system_types_to_drop2)]
        # Filter out rows based on system types to drop (same for Air Flow Rate and Fan Power)
        sva_result3 = sva_result[~sva_result['SYSTEM_TYPE'].isin(system_types_to_drop3)]
        # Filter out rows based on system types to drop (Heating Capacity)
        sva_result4 = sva_result[~sva_result['SYSTEM_TYPE'].isin(system_types_to_drop4)]

        sva_result1 = sva_result1[~sva_result1.apply(lambda row: row.astype(str).str.contains('FILE-')).any(axis=1)]
        sva_result2 = sva_result2[~sva_result2.apply(lambda row: row.astype(str).str.contains('FILE-')).any(axis=1)]
        sva_result3 = sva_result3[~sva_result3.apply(lambda row: row.astype(str).str.contains('FILE-')).any(axis=1)]
        sva_result4 = sva_result4[~sva_result4.apply(lambda row: row.astype(str).str.contains('FILE-')).any(axis=1)]

        # Assuming sva_result is your DataFrame
        sva_result1['CAPACITY'] = pd.to_numeric(sva_result1['CAPACITY'])
        sva_result1['SYS_COOLING_EIR'] = pd.to_numeric(sva_result1['SYS_COOLING_EIR'])
        sva_result1['SYS_COOLING_CAPACITY'] = pd.to_numeric(sva_result1['SYS_COOLING_CAPACITY'])
        sva_result1['SYS_HEATING_EIR'] = pd.to_numeric(sva_result1['SYS_HEATING_EIR'])
        sva_result1['SYS_HEATING_CAPACITY'] = pd.to_numeric(sva_result1['SYS_HEATING_CAPACITY'])
        sva_result1['ZONE_COOLING_CAPACITY'] = pd.to_numeric(sva_result1['ZONE_COOLING_CAPACITY'])
        sva_result1['ZONE_HEATING_CAPACITY'] = pd.to_numeric(sva_result1['ZONE_HEATING_CAPACITY'])
        sva_result1['ZONE_MULT'] = pd.to_numeric(sva_result1['ZONE_MULT'])

        # Assuming sva_result is your DataFrame
        sva_result2['CAPACITY'] = pd.to_numeric(sva_result2['CAPACITY'])
        sva_result2['SYS_COOLING_EIR'] = pd.to_numeric(sva_result2['SYS_COOLING_EIR'])
        sva_result2['SYS_COOLING_CAPACITY'] = pd.to_numeric(sva_result2['SYS_COOLING_CAPACITY'])
        sva_result2['SYS_HEATING_EIR'] = pd.to_numeric(sva_result2['SYS_HEATING_EIR'])
        sva_result2['SYS_HEATING_CAPACITY'] = pd.to_numeric(sva_result2['SYS_HEATING_CAPACITY'])
        sva_result2['ZONE_COOLING_CAPACITY'] = pd.to_numeric(sva_result2['ZONE_COOLING_CAPACITY'])
        sva_result2['ZONE_HEATING_CAPACITY'] = pd.to_numeric(sva_result2['ZONE_HEATING_CAPACITY'])
        sva_result2['ZONE_MULT'] = pd.to_numeric(sva_result2['ZONE_MULT'])

        # Assuming sva_result is your DataFrame
        sva_result3['CAPACITY'] = pd.to_numeric(sva_result3['CAPACITY'])
        sva_result3['SYS_COOLING_EIR'] = pd.to_numeric(sva_result3['SYS_COOLING_EIR'])
        sva_result3['SYS_COOLING_CAPACITY'] = pd.to_numeric(sva_result3['SYS_COOLING_CAPACITY'])
        sva_result3['SYS_HEATING_EIR'] = pd.to_numeric(sva_result3['SYS_HEATING_EIR'])
        sva_result3['SYS_HEATING_CAPACITY'] = pd.to_numeric(sva_result3['SYS_HEATING_CAPACITY'])
        sva_result3['ZONE_COOLING_CAPACITY'] = pd.to_numeric(sva_result3['ZONE_COOLING_CAPACITY'])
        sva_result3['ZONE_HEATING_CAPACITY'] = pd.to_numeric(sva_result3['ZONE_HEATING_CAPACITY'])
        sva_result3['ZONE_MULT'] = pd.to_numeric(sva_result3['ZONE_MULT'])
        sva_result3['SUPPLY_FLOW'] = pd.to_numeric(sva_result3['SUPPLY_FLOW'])
        sva_result3['FAN'] = pd.to_numeric(sva_result3['FAN'])
        sva_result3['POWER_DEMAND'] = pd.to_numeric(sva_result3['POWER_DEMAND'])

        # Assuming sva_result is your DataFrame
        sva_result4['CAPACITY'] = pd.to_numeric(sva_result4['CAPACITY'])
        sva_result4['SYS_COOLING_EIR'] = pd.to_numeric(sva_result4['SYS_COOLING_EIR'])
        sva_result4['SYS_COOLING_CAPACITY'] = pd.to_numeric(sva_result4['SYS_COOLING_CAPACITY'])
        sva_result4['SYS_HEATING_EIR'] = pd.to_numeric(sva_result4['SYS_HEATING_EIR'])
        sva_result4['SYS_HEATING_CAPACITY'] = pd.to_numeric(sva_result4['SYS_HEATING_CAPACITY'])
        sva_result4['ZONE_COOLING_CAPACITY'] = pd.to_numeric(sva_result4['ZONE_COOLING_CAPACITY'])
        sva_result4['ZONE_HEATING_CAPACITY'] = pd.to_numeric(sva_result4['ZONE_HEATING_CAPACITY'])
        sva_result4['ZONE_MULT'] = pd.to_numeric(sva_result4['ZONE_MULT'])

        sva_result1['COOLING_ELECTRIC_POWER'] = \
            ((sva_result1['ZONE_COOLING_CAPACITY'] * sva_result1['SYS_COOLING_EIR']) if sva_result1['SYSTEM_TYPE'].isin(['PTAC', 'HP', 'FC']).any() \
                else sva_result1['SYS_COOLING_CAPACITY'] * sva_result1['SYS_COOLING_EIR'])
        
        sva_result1['COOLING_ELECTRIC_POWER'] = pd.to_numeric(sva_result1['COOLING_ELECTRIC_POWER'])

        # Apply the conditions directly to create the new column
        sva_result1['NEW_COOLING_ELECTRIC_POWER'] = \
            (sva_result1['ZONE_COOLING_CAPACITY'] if sva_result1['SYSTEM_TYPE'].isin(['PTAC', 'HP', 'FC']).any() else sva_result1['SYS_COOLING_CAPACITY']) / ((sva_result1['ZONE_COOLING_CAPACITY'] * sva_result1['SYS_COOLING_EIR']) if sva_result1['SYSTEM_TYPE'].isin(['PTAC', 'HP', 'FC']).any() \
                else sva_result1['SYS_COOLING_CAPACITY'] * sva_result1['SYS_COOLING_EIR'])
        # sva_result1['COOLING_ELECTRIC_POWER'] / \
        sva_result2['NEW_COOLING_CAPACITY'] = \
            (sva_result2['ZONE_COOLING_CAPACITY'] * sva_result2['ZONE_MULT'] if sva_result2['SYSTEM_TYPE'].isin(['FC', 'PTAC', 'HP']).any() else sva_result2['SYS_COOLING_CAPACITY'])
        

        sva_result3['AIR_FLOW_RATE'] = (sva_result3['SUPPLY_FLOW'] * sva_result3['ZONE_MULT'])
        sva_result3['FAN_POWER'] = \
            (sva_result3['FAN'] * sva_result3['ZONE_MULT'] if sva_result3['SYSTEM_TYPE'].isin(['FC', 'PTAC', 'HP', 'UHT', 'UVT']).any() else sva_result3['POWER_DEMAND'])
            
        sva_result4['NEW_HEATING_CAPACITY'] = \
            (sva_result4['ZONE_HEATING_CAPACITY'] * sva_result4['ZONE_MULT'] if sva_result4['SYSTEM_TYPE'].isin(['FC', 'FPH', 'PTAC', 'HP', 'UHT', 'UVT']).any() else sva_result4['SYS_HEATING_CAPACITY'])

        # print("\n**COOLING ELECTRIC POWER(KW), COOLING EIR(BTU/BTU), COOLING_CAPACITY(KBTU/HR), AIR_FLOW_RATE(CFM), FAN POWER(KW), HEATING CAPACITY(KBTU/HR)**")
        # print("********",x[0],"********\n")
        total_electric_power = sva_result1['COOLING_ELECTRIC_POWER'].sum()
        matching_sum_0 = sva_result1.groupby('SYSTEM_TYPE')['COOLING_ELECTRIC_POWER'].sum() # 1
        # Print SYSTEM_TYPE and corresponding matching sum
        # for system_type, sum_value in matching_sum_0.items():
        #      print(f"SYSTEM_TYPE: {system_type}, Cooling_Electric_Power: {sum_value}")

        System_Cooling_eir = sva_result1['NEW_COOLING_ELECTRIC_POWER'].sum()
        # matching_sum_0 = sva_result1.groupby('SYSTEM_TYPE')['COOLING_ELECTRIC_POWER'].sum() # 1
        matching_sum_0_ = sva_result1.groupby('SYSTEM_TYPE')['NEW_COOLING_ELECTRIC_POWER'].sum() # 1
        # Print SYSTEM_TYPE and corresponding matching sum
        # for system_type, sum_value in matching_sum_0_.items():
        #     if sum_value != 0:
        #         print(f"SYSTEM_TYPE: {system_type}, Cooling_EIR: {1/sum_value}")
        #     else:
        #         print(f"SYSTEM_TYPE: {system_type}, Cooling_EIR: {sum_value}")
    
        cooling_capacity = sva_result2['NEW_COOLING_CAPACITY'].sum() # 2
        matching_sum_2 = sva_result2.groupby('SYSTEM_TYPE')['NEW_COOLING_CAPACITY'].sum() # 2
        # Print SYSTEM_TYPE and corresponding matching sum
        # for system_type, sum_value in matching_sum_2.items():
        #     print(f"SYSTEM_TYPE: {system_type}, Cooling_Capacity: {sum_value}")

        Air_Flow_rate = sva_result3['AIR_FLOW_RATE'].sum() # 3
        matching_sum_3 = sva_result3.groupby('SYSTEM_TYPE')['AIR_FLOW_RATE'].sum() # 3
        # Print SYSTEM_TYPE and corresponding matching sum
        # for system_type, sum_value in matching_sum_3.items():
        #     print(f"SYSTEM_TYPE: {system_type}, AFR: {sum_value}")

        Fan_Power = sva_result3['FAN_POWER'].sum() # 3
        matching_sum_3_ = sva_result3.groupby('SYSTEM_TYPE')['FAN_POWER'].sum() # 3
        # Print SYSTEM_TYPE and corresponding matching sum
        # for system_type, sum_value in matching_sum_3_.items():
        #     print(f"SYSTEM_TYPE: {system_type}, FAN POWER: {sum_value}")

        heating_capacity = sva_result4['NEW_HEATING_CAPACITY'].sum() # 4
        matching_sum_4 = sva_result4.groupby('SYSTEM_TYPE')['NEW_HEATING_CAPACITY'].sum() # 4
        # Print SYSTEM_TYPE and corresponding matching sum
        # for system_type, sum_value in matching_sum_4.items():
        #     print(f"SYSTEM_TYPE: {system_type}, HEAT CAPACITY: {sum_value}")

        sva_result1.drop(sva_result1.columns[[-4, -6]], axis=1, inplace=True)
        sva_result2.drop(sva_result2.columns[[-3, -5]], axis=1, inplace=True)
        sva_result3.drop(sva_result3.columns[[-4, -6]], axis=1, inplace=True)
        sva_result4.drop(sva_result4.columns[[-3, -5]], axis=1, inplace=True)
        # print(Air_Flow_rate)

    return [cooling_capacity, System_Cooling_eir, Air_Flow_rate, Fan_Power, heating_capacity]


# function to get all values from lvb, lvd, _zone, and Summary files. having parameters combined_data and takes input path.
def get_all_calculated_values(combined_data, inputPath):
    # creating empty list to store new column in combined data.
    FTAbove = []
    FCTotal = []
    FUTotal = []
    FTBelow = []
    WallTAbobe = []
    powerLight = []
    EquipTotal = []
    LscVal = []
    EnergyOutcome = []
    EnergyOutcome_Therm = []
    Ratio_WWR = []
    eflh = []
    totload = []
    shgc = []

    # this is user input, it is considered as blank for now.
    AspectRatio = [None]*len(combined_data)
    POrientation = [None]*len(combined_data)

    # create empty dataframe to store wallArea and wallU values.
    wallArea = pd.DataFrame()
    wallU = pd.DataFrame()
    windowArea = pd.DataFrame()
    windowU = pd.DataFrame()

    # Loop through subfolders and files
    subfolders = os.listdir(inputPath)
    for subfolder in subfolders:
        print(subfolder)
        path = os.path.join(inputPath, subfolder)

        # store _lvd.csv data in lvbfiles varibales, like this-
        lvbfiles = gb.glob(f'{path}/*_lvb.csv', recursive=True)
        lvdfiles = gb.glob(f'{path}/*_lvd.csv', recursive=True)
        zonefiles = gb.glob(f'{path}/*_Zone.csv', recursive=True)
        lvd_summary = gb.glob(f'{path}/*_lvd_Summary.csv', recursive=True)
        lscfiles = gb.glob(f'{path}/*_lsc.csv', recursive=True)
        bepufiles = gb.glob(f'{path}/*_bepu.csv', recursive=True)
        svafiles = gb.glob(f'{path}/*_sva.csv', recursive=True)
        psefiles = gb.glob(f'{path}/*_pse.csv', recursive=True)
        shgcfiles = gb.glob(f'{path}/*_shgc.csv', recursive=True)

        print(bepufiles)

        # iterate in each csvfiles variables at a time using zip function
        for lvbfile, lvdfile, zonefile, summary, lscfile, bepufile, svafile, psefile, shgcfile in zip(lvbfiles, lvdfiles, zonefiles, lvd_summary, lscfiles, bepufiles, svafiles, psefiles, shgcfiles):
            # store all csvs info. in dataframe
            lvb_data = pd.read_csv(lvbfile)
            lvd_data = pd.read_csv(lvdfile)
            zone_data = pd.read_csv(zonefile)
            summary_data = pd.read_csv(summary)
            lsc_data = pd.read_csv(lscfile)
            bepu_data = pd.read_csv(bepufile)
            sva_data = pd.read_csv(svafile)
            pse_data = pd.read_csv(psefile)
            shgc_data = pd.read_csv(shgcfile)

            print("Raw data in " + path)
            print(lvb_data)
            print(lvd_data)
            print(zone_data)
            print(summary_data)
            print(lsc_data)
            print(bepu_data)
            print(sva_data)
            print(pse_data)
            print(shgc_data)
            exit()

            LscVal.append(lsc_data) # all column of lsc_csv in combined_data.
            Totalenergy = bepu_data['TOTAL-BEPU'].sum()
            EnergyOutcome.append(Totalenergy)

            shgc_value = shgc_data['SHADING-COEF'].sum()
            shgc.append(shgc_value)

            index_value = summary_data[summary_data['AZIMUTH'] == 'ALL WALLS'].index[0]
            output_value = summary_data.loc[index_value, 'WINDOW+WALL(AREA)(SQFT)']

            wind = summary_data.loc[summary_data['AZIMUTH'] == 'ALL WALLS', 'WINDOW(AREA)(SQFT)'].values[0]
            wind_wall = summary_data.loc[summary_data['AZIMUTH'] == 'ALL WALLS', 'WINDOW+WALL(AREA)(SQFT)'].values[0]
            ratio = wind/wind_wall
            Ratio_WWR.append(ratio)
            ratioPSE = pse_data["TOTAL"].iloc[-2] / pse_data["TOTAL"].iloc[-1]
            eflh.append(ratioPSE)
            totload.append(pse_data["TOTAL"].iloc[-1])

            # it will store total area of all projects in wallTAbobe list.
            WallTAbobe.append(output_value)

            ans1 = lvb_data[['SPACE']]
            ans2 = lvd_data[['SPACE', 'Grade-Expression']].drop_duplicates()
            l1 = ans1['SPACE'].to_list()
            l2 = ans2['SPACE'].to_list()
            k = ans2['Grade-Expression'].to_list()
            for i in range(0, len(l1) - len(k)):
                k.append('BG')
            l3 = []
            for i in range(0,len(l1)):
                found = False
                for j in range(0,len(l2)):
                    if l1[i] in l2[j]:
                        l3.append(k[j])
                        found = True
                        break
                if not found:
                    l3.append('AG')
            lvb_data['Grade-Expression'] = l3

            # Area Total
            areaTot = []
            for i in range(0, len(l3)):
                areaTot.append(lvb_data['AREA(SQFT)'][i]*lvb_data['SPACE*FLOOR'][i])
            lvb_data['Area-Total(SQFT)'] = areaTot

            # Lighting Power Total
            lightTot = []
            totLightPower = 0
            for j in range(0, len(l3)):
                lightTot.append(lvb_data['LIGHTS'][j]*lvb_data['Area-Total(SQFT)'][j])
                totLightPower = totLightPower + lightTot[j]
            lvb_data['Lighting-Power-Total(WATT)'] = lightTot
            powerLight.append(totLightPower)

            # Equipment Total
            equipTot = []
            totalEquip = 0
            for k in range(0, len(l3)):
                equipTot.append(lvb_data['EQUIP'][k]*lvb_data['Area-Total(SQFT)'][k])
                totalEquip = totalEquip + equipTot[k]
            lvb_data['Equipment-Total(WATT)'] = equipTot
            EquipTotal.append(totalEquip)

            BGArea = 0
            AGArea = 0
            for ele in range(0,len(l3)):
                if l3[ele] == 'AG':
                    AGArea = AGArea + lvb_data['AREA(SQFT)'][ele]
                else:
                    BGArea = BGArea + lvb_data['AREA(SQFT)'][ele]
            FTAbove.append(AGArea)
            FTBelow.append(BGArea)
            
            if 'BUILDING-TYPE' in lvb_data.columns and lvb_data['BUILDING-TYPE'].dtype == 'object':
                conditioned1 = lvb_data[~lvb_data['BUILDING-TYPE'].str.contains('U')]['AREA(SQFT)'].sum()
                unconditioned1 = lvb_data[lvb_data['BUILDING-TYPE'].str.contains('U')]['AREA(SQFT)'].sum()
            else:
                conditioned1 = 0
                unconditioned1 = 0

            FCTotal.append(conditioned1)
            FUTotal.append(unconditioned1)
    
            ### For Summary Data Information ###
            summary_data = summary_data.drop('RUNNAME', axis = 1)
            data = summary_data.to_dict()
            data1 = data['AZIMUTH']
            data2 = data['AVERAGE(U-VALUE/WINDOWS)(BTU/HR-SQFT-F)']
            data3 = data['AVERAGE(U-VALUE/WALLS)(BTU/HR-SQFT-F)']
            data4 = data['AVERAGE U-VALUE(WALLS+WINDOWS)(BTU/HR-SQFT-F)']
            data5 = data['WINDOW(AREA)(SQFT)']
            data6 = data['WALL(AREA)(SQFT)']
            data7 = data['WINDOW+WALL(AREA)(SQFT)']
            azimuth_column = {0: 'NORTH', 1: 'SOUTH', 2: 'EAST', 3: 'WEST', 4: 'NORTH-EAST', 5: 'NORTH-WEST', 6: 'SOUTH-EAST',
                7: 'SOUTH-WEST', 8: 'ROOF', 9: 'ALL WALLS', 10: 'WALLS+ROOFS', 11: 'UNDERGRND', 12: 'BUILDING'}
            wallArea_temp = pd.DataFrame(columns=azimuth_column.values())
            for key in data1:
                header = data1[key]
                if header in azimuth_column.values():
                    col_index = list(azimuth_column.values()).index(header)
                    wallArea_temp[header] = [data6[key]]
                else:
                    wallArea_temp[header] = ''
            wallArea_temp.columns = ['N-Wall-Area(SQFT)', 'S-Wall-Area(SQFT)', 'E-Wall-Area(SQFT)', 'W-Wall-Area(SQFT)', 'NE-Wall-Area(SQFT)', 'NW-Wall-Area(SQFT)', 'SE-Wall-Area(SQFT)',
                                'SW-Wall-Area(SQFT)', 'ROOF-AREA(SQFT)', 'ALL WALLS-Wall-AREA(SQFT)', 'WALLS+ROOFS-AREA(SQFT)', 'UNDERGRND-Wall-AREA(SQFT)', 'BUILDING-Wall-AREA(SQFT)']
            wallArea = pd.concat([wallArea, wallArea_temp], ignore_index=True)

            wallU_temp = pd.DataFrame(columns=azimuth_column.values())
            for key in data1:
                header = data1[key]
                if header in azimuth_column.values():
                    col_index = list(azimuth_column.values()).index(header)
                    wallU_temp[header] = [data3[key]]
                else:
                    wallU_temp[header] = ''
            wallU_temp.columns = ['N-Wall-U-Value(BTU/HR-SQFT-F)', 'S-Wall-U-Value(BTU/HR-SQFT-F)', 'E-Wall-U-Value(BTU/HR-SQFT-F)', 'W-Wall-U-Value(BTU/HR-SQFT-F)', 'NE-Wall-U-Value(BTU/HR-SQFT-F)', 'NW-Wall-U-Value(BTU/HR-SQFT-F)', 'SE-Wall-U-Value(BTU/HR-SQFT-F)',
                                'SW-Wall-U-Value(BTU/HR-SQFT-F)', 'ROOF-U-Value(BTU/HR-SQFT-F)', 'ALL WALLS-Wall-U-Value(BTU/HR-SQFT-F)', 'WALLS+ROOFS-U-Value(BTU/HR-SQFT-F)', 'UNDERGRND-Wall-U-Value(BTU/HR-SQFT-F)', 'BUILDING-Wall-U-Value(BTU/HR-SQFT-F)']
            wallU = pd.concat([wallU, wallU_temp], ignore_index=True)

            windowArea_temp = pd.DataFrame(columns=azimuth_column.values())
            for key in data1:
                header = data1[key]
                if header in azimuth_column.values():
                    col_index = list(azimuth_column.values()).index(header)
                    windowArea_temp[header] = [data5[key]]
                else:
                    windowArea_temp[header] = ''
            windowArea_temp.columns = ['N-Window-Area(SQFT)', 'S-Window-Area(SQFT)', 'E-Window-Area(SQFT)', 'W-Window-Area(SQFT)', 'NE-Window-Area(SQFT)', 'NW-Window-Area(SQFT)', 'SE-Window-Area(SQFT)',
                                'SW-Window-Area(SQFT)', 'ROOF-Window-Area(SQFT)', 'ALL WALLS-Window-AREA(SQFT)', 'WALLS+ROOFS-Window-AREA(SQFT)', 'UNDERGRND-Window-AREA(SQFT)', 'BUILDING-Window-AREA(SQFT)']
            windowArea = pd.concat([windowArea, windowArea_temp], ignore_index=True)

            windowU_temp = pd.DataFrame(columns=azimuth_column.values())
            for key in data1:
                header = data1[key]
                if header in azimuth_column.values():
                    col_index = list(azimuth_column.values()).index(header)
                    windowU_temp[header] = [data2[key]]
                else:
                    windowU_temp[header] = ''
            windowU_temp.columns = ['N-Window-U-Value(BTU/HR-SQFT-F)', 'S-Window-U-Value(BTU/HR-SQFT-F)', 'E-Window-U-Value(BTU/HR-SQFT-F)', 'W-Window-U-Value(BTU/HR-SQFT-F)', 'NE-Window-U-Value(BTU/HR-SQFT-F)', 'NW-Window-U-Value(BTU/HR-SQFT-F)', 'SE-Window-U-Value(BTU/HR-SQFT-F)',
                                'SW-Window-U-Value(BTU/HR-SQFT-F)', 'ROOF-Window-U-Value(BTU/HR-SQFT-F)', 'ALL WALLS-Window-U-Value(BTU/HR-SQFT-F)', 'WALLS+ROOFS-Window-U-Value(BTU/HR-SQFT-F)', 'UNDERGRND-Window-U-Value(BTU/HR-SQFT-F)', 'BUILDING-Window-U-Value(BTU/HR-SQFT-F)']
            windowU = pd.concat([windowU, windowU_temp], ignore_index=True)

    combined_data['Floor-Total-Above-Grade(SQFT)'] = FTAbove ### USE in case of- Wall/Floor ratio and window/Floor ratio ### 
    combined_data['Floor-Total-Below-Grade(SQFT)'] = FTBelow
    combined_data['Floor-Total-Conditioned-Grade(SQFT)'] = FCTotal
    combined_data['Floor-Total-UnConditioned-Grade(SQFT)'] = FUTotal
    combined_data['Wall-Total-Above-Grade(SQFT)'] = WallTAbobe  ### USE in case of- Wall/Floor ratio ### 
    combined_data['Power Lighting Total(W/SQFT)'] = powerLight
    combined_data['Equipment-Total(W/SQFT)'] = EquipTotal

    # Reset the indexes of all the DataFrames
    combined_data.reset_index(drop=True, inplace=True)
    wallArea.reset_index(drop=True, inplace=True)
    wallU.reset_index(drop=True, inplace=True)
    windowArea.reset_index(drop=True, inplace=True)
    windowU.reset_index(drop=True, inplace=True)
    lsc_data.reset_index(drop=True, inplace=True)

    # Concatenate the DataFrames along the columns axis
    result = pd.concat([combined_data, wallArea, wallU, windowArea, windowU], axis=1)
    concatenated_lsc_data = pd.concat(LscVal, ignore_index=True)
    
    # Concatenate the DataFrames along the columns axis
    result = pd.concat([result, concatenated_lsc_data], axis=1)
    result['WWR'] = Ratio_WWR
    result['EFLH'] = eflh
    result = result.drop('TOTAL-LOAD(KW)', axis=1)
    result['TOTAL-LOAD(KW)'] = totload
    result['SC'] = shgc
    result['Energy_Outcome(KWH)'] = EnergyOutcome
    # Reset the indexes of all the DataFrames
    result.reset_index(drop=True, inplace=True)
    print("\nAll CSVs are generated in the respective folder!")

    return result
    # except Exception as e:
    #     print(f"Moving to next")