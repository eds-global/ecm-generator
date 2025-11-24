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
            other_str = []  # List to store lines preceding the 'MBTU' lines
            prev_line = None  # Initialize variable to store the previous line
            for line in lvc_rpt:
                if prev_line:
                    if ('.' in line):
                        lvc_str.append(line)  # Append the current line
                        other_str.append(prev_line)  # Store the previous line
                # Store the current line as the previous line for the next iteration
                prev_line = line
            
            # this list will store the bepu_list   
            result = [] 
            # Iterate through each line in bepu_rpt
            for line in lvc_str:
                lvc_list = []
                # Split the line by whitespace and store the result in splitter
                splitter = line.split()
                # Join the first part of the splitter except the last 13 elements and store it as space_name
                space_name = " ".join(splitter[:-130])
                # Add space_name as the first element of bepu_list
                lvc_list=splitter[-130:]
                # Add space_name as the first element of bepu_list
                lvc_list.insert(0,space_name)
                # append bepu_list to result
                result.append(lvc_list)
            # store result to dataframe
            lvc_df = pd.DataFrame(result)
            
            return lvc_df

    except Exception as e:
        columns = ['RUNNAME', 'BEPU-SOURCE', 'BEPU-UNIT', 'LIGHTS', 'TASK-LIGHTS', 'MISQ-EQUIP', 'SPACE-HEATING',
                                'SPACE-COOLING', 'HEAT-REJECT', 'PUMPS & AUX', 'VENT FANS', 'REFRING-DISPLAY',
                                'HT-PUMP-SUPPLEMENT', 'DOMEST-HOT-WTR', 'EXT-USAGE', 'TOTAL-BEPU']
        return pd.DataFrame(columns=columns)

def get_SHGC_report(name, path=None):
    try:
        with open(name, 'r') as f:
            flist = f.readlines()

        # Extract lines between "Glass Types" and "Window Layers"
        start, end = None, None
        for idx, line in enumerate(flist):
            if 'Glass Types' in line:
                start = idx
            if 'Window Layers' in line:
                end = idx
                break

        if start is None or end is None:
            raise ValueError("Markers 'Glass Types' or 'Window Layers' not found.")

        glass_section = flist[start:end]

        glass_data = []
        current_glass = None
        glass_info = {}

        for line in glass_section:
            line = line.strip()

            # Detect new glass type entry
            if line.startswith('"') and '= GLASS-TYPE' in line:
                if current_glass and glass_info:
                    glass_info['GLASS-TYPE'] = current_glass
                    glass_data.append(glass_info)

                current_glass = line.split('=')[0].strip().strip('"')
                glass_info = {}

            elif '=' in line:
                key, value = map(str.strip, line.split('='))
                glass_info[key] = value

        # Append the last one
        if current_glass and glass_info:
            glass_info['GLASS-TYPE'] = current_glass
            glass_data.append(glass_info)

        # Convert to DataFrame
        df = pd.DataFrame(glass_data)

        # Optional: Reorder columns
        cols = ['GLASS-TYPE'] + [col for col in df.columns if col != 'GLASS-TYPE']
        df = df[cols]

        return df

    except Exception as e:
        print(f"Error: {e}")
        return pd.DataFrame()  # return empty df on failure
