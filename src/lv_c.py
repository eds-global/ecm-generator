import os
import pandas as pd
import warnings

warnings.filterwarnings("ignore")

def get_LVC_report(name, path):
    try:
        with open(name) as f:
            # Read all lines from the file and store them in a list named flist
            flist = f.readlines()

            lvc_count = [] 
            # Iterate through each line in flist along with its line number
            for num, line in enumerate(flist, 0):
                if 'LV-C' in line:
                    lvc_count.append(num)
                if 'LV-D' in line:
                    numend = num
            numstart = lvc_count[0] 
            lvc_rpt = flist[numstart:numend]

            lvc_str = []  # List to store lines containing 'MBTU'
            for line in lvc_rpt:
                if ('.' in line and 'Win' in line):
                    lvc_str.append(line)  # Append the current line
            
            # this list will store the bepu_list   
            result = [] 
            # Iterate through each line in bepu_rpt
            for line in lvc_str:
                lvc_list = []
                # Split the line by whitespace and store the result in splitter
                splitter = line.split()
                # Join the first part of the splitter except the last 13 elements and store it as space_name
                space_name = " ".join(splitter[0:])
                # Add space_name as the first element of bepu_list
                lvc_list=splitter[0:]
                # Add space_name as the first element of bepu_list
                lvc_list.insert(0,space_name)
                # append bepu_list to result
                result.append(lvc_list)
            # store result to dataframe
            lvc_df = pd.DataFrame(result)
            lvc_df = lvc_df.iloc[:, 1:]
            # Drop rows where the last column is blank, whitespace, or NaN
            last_col = lvc_df.columns[-1]
            lvc_df = lvc_df[~(lvc_df[last_col].isna() | lvc_df[last_col].astype(str).str.strip().eq(""))]
            # Merge all columns except the last 10
            merged_col = lvc_df.iloc[:, :-10].astype(str).apply(' '.join, axis=1)
            last_10 = lvc_df.iloc[:, -10:]
            lvc_df = pd.concat([merged_col.rename('Window'), last_10], axis=1)
            lvc_df.columns = ["WINDOW", "MULTIPLIER", "GLASS AREA (SQFT)", "GLASS WIDTH (FT)", "GLASS HEIGHT (FT)", "SET-BACK (FT)", "NUMBER OF PANES", "CENTER-OF-GLASS U-VALUE (BTU/HR-SQFT-F)", "GLASS SHADING COEFF", "GLASS VISIBLE TRANS", "GLASS SOLAR TRANS"]
            # print(lvc_df)
            return lvc_df

    except Exception as e:
        columns = ["WINDOW", "MULTIPLIER", "GLASS AREA (SQFT)", "GLASS WIDTH (FT)", "GLASS HEIGHT (FT)", "SET-BACK (FT)", "NUMBER OF PANES", "CENTER-OF-GLASS U-VALUE (BTU/HR-SQFT-F)", "GLASS SHADING COEFF", "GLASS VISIBLE TRANS", "GLASS SOLAR TRANS"]
        return pd.DataFrame(columns=columns)