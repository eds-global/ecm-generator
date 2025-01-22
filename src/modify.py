import os
import streamlit as st
import tempfile
import zipfile
from src.lpd_modify import _lpd
from src.lpd_modify import _epd
from src.lpd_modify import _orient
from src.lpd_modify import _wwr

def update_inp_file(uploaded_file, modified_values, idx):
    """
    Update the uploaded INP file based on selected ECMs and return the content for download.
    """
    # Extract ECM values from the modified values
    lpd = modified_values.get("LPD", None)
    epd = modified_values.get("EPD", None)
    wwr = modified_values.get("WWR", None)
    orient = modified_values.get("Orient", None)
    wall = modified_values.get("Wall-Type", None)
    roof = modified_values.get("Roof-Type", None)
    window = modified_values.get("Window-Type", None)

    # Check if all inputs are None and return a message if true
    if all(value is None for value in [lpd, epd, wwr, orient, wall, roof, window]):
        st.info("❌ Invalid input! Please enter some text to modify.")
        return None, None

    if uploaded_file is not None:
        # try:
            # Create a temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save the uploaded file temporarily
            inp_path = os.path.join(temp_dir, uploaded_file.name)
            with open(inp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            inp_path = inp_path.replace('\n', '\r\n')

            # Process the ECM updates
            if lpd is not None or epd is not None or wwr is not None or orient is not None or wall is not None or roof is not None or window is not None:
                _data = inp_path  # Start with the original data
                _data = _epd.perging_data_weekly(_data, epd)
                _data = _lpd.perging_data_annual(_data, lpd)
                # _data = _wwr.process_all_inp_files_in_folder(_data, wwr)  # Assuming 'df' is passed here in your original code
                _data = _orient.getOrientation(_data, orient)

                # Generate the updated INP file
                base_name, ext = os.path.splitext(uploaded_file.name)
                updated_file_name = f"{base_name}_ECM_Set_{idx}{ext}"
                updated_file_path = os.path.join(temp_dir, updated_file_name)

                with open(updated_file_path, 'w', newline='\r\n') as file:
                    file.writelines(_data)

                # Read the updated file content
                with open(updated_file_path, 'rb') as file:
                    file_content = file.read()

                return file_content, updated_file_name  # Return the file content and name

        # except Exception as e:
        #     raise Exception(f"An error occurred while updating INP file: {e}")



















                    # if wwr is not None:
                    #     polygon_df = _wwr.extract_polygons(_data, wwr)
                    #     df = _wwr.process_inp_file(_data)
                        
                    #     df["Next-Column"] = df.apply(get_next_column, axis=1)
                    #     df["Next-Column"] = df.apply(get_next_column, axis=1)
                    #     df["EW-LOC_Num"] = df["LOCATION"].str.extract(r'(V\d+)')
                    #     df["Next-Column_Num"] = df["Next-Column"].str.extract(r'(V\d+)')
                    #     df["Diff"] = df["Next-Column_Num"] + " - " + df["EW-LOC_Num"]
                    #     df["Diff"] = df["Diff"].fillna("")
                    #     df = df.drop(columns=["EW-LOC_Num", "Next-Column_Num"])
                    #     df['EXTERIOR-WINDOW'] = df['EXTERIOR-WALL'].apply(create_ext_win)
                    #     columns = list(df.columns)
                    #     ext_wall_index = columns.index('EXTERIOR-WALL')
                    #     columns.insert(ext_wall_index + 1, columns.pop(columns.index('EXTERIOR-WINDOW')))
                    #     df = df[columns]

                    #     df = pd.merge(df, polygon_df, left_on='POLYGON', right_on='Polygon', how='inner')
                    #     df = df.drop(columns=['Polygon'])

                    #     df["Cordinate"] = df.apply(calculate_corr, axis=1)

                    #     columns_to_remove = [f'V{i}' for i in range(1, 101)]
                    #     df = df.drop(columns=[col for col in columns_to_remove if col in df.columns])

                    #     df['D'] = df['Cordinate'].apply(lambda x: round(calculate_distance(x), 2) if x is not None else None)

                    #     # Generate percentages from 10% to 90%
                    #     df['SPACE-HEIGHT'] = pd.to_numeric(df['SPACE-HEIGHT'], errors='coerce')
                    #     percentages = [i / 100 for i in range(5, 100, 5) if i == 5 or i % 10 == 0]
                    #     percentages_SH = [i / 100 for i in range(5, 100, 5) if i == 5 or i % 10 == 0]

                    #     for percent in percentages:
                    #         column_name = f"D{int(percent * 100)}%"  # Column names will be D10%, D20%, ..., D90%
                    #         df[column_name] = df['D'] * percent

                    #     for percent in percentages_SH:
                    #         column_name = f"SPACE-HEIGHT{int(percent * 100)}%"  # Column names will be D10%, D20%, ..., D90%
                    #         df[column_name] = df['SPACE-HEIGHT'] * percent

                    #     percentage_columns = [f"D{int(percent * 100)}%" for percent in percentages]
                    #     percentage_columns_SH = [f"SPACE-HEIGHT{int(percent * 100)}%" for percent in percentages_SH]
                    #     df[percentage_columns] = df[percentage_columns].round(2)
                    #     df[percentage_columns_SH] = df[percentage_columns_SH].round(2)

                    #     df['X'] = df['D5%']
                    #     df['Y'] = df['SPACE-HEIGHT5%']

                    #     for i, factor in enumerate([0.4, 0.5, 0.6, 0.7, 0.8, 0.9], start=1):
                    #         df[f'HEIGHT{i}'] = factor * df['SPACE-HEIGHT']
                    #         df[f'WIDTH{i}'] = factor * df['D']

                    #     df['EXTERIOR-WINDOW'] = df['EXTERIOR-WALL'].apply(create_ext_win)
                    #     columns = list(df.columns)
                    #     ext_wall_index = columns.index('EXTERIOR-WALL')
                    #     columns.insert(ext_wall_index + 1, columns.pop(columns.index('EXTERIOR-WINDOW')))
                    #     df = df[columns]
                    #     _data = purge_windows.process_all_inp_files_in_folder(inp_path, df)