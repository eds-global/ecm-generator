import os
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

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