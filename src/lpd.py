import os
import streamlit as st
import tempfile
from src.lpd_modify import _lpd
from src.lpd_modify import _epd

def update_inp_file(uploaded_file, modified_values, idx):
    # Extract LPD and EPD
    lpd = modified_values.get("LPD", None)
    epd = modified_values.get("EPD", None)
    wwr = modified_values.get("WWR", None)
    orient = modified_values.get("Orient", None)
    wall = modified_values.get("Wall-Type", None)
    roof = modified_values.get("Roof-Type", None)
    window = modified_values.get("Window-Type", None)
   
    if uploaded_file is not None:
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save the uploaded file temporarily
                inp_path = os.path.join(temp_dir, uploaded_file.name)
                
                with open(inp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
 
                inp_path = inp_path.replace('\n', '\r\n')
                if lpd is not None or epd is not None or wwr is not None or orient is not None or wall is not None or roof is not None or window is not None:
                    if lpd is not None:
                        _data = _lpd.perging_data_annual(inp_path, lpd)
                    if epd is not None:
                        _data = _epd.perging_data_weekly(_data, epd)
                    if wwr is not None:
                        _data = _wwr.perging_data_annual(_data, wwr)
                    if orient is not None:
                        _data = _orient.perging_data_weekly(_data, orient)
        
                # Create the updated INP file
                base_name, ext = os.path.splitext(uploaded_file.name)
                updated_file_name = f"{base_name}_ECM_Set_{idx}{ext}"
                updated_file_path = os.path.join(temp_dir, updated_file_name)

                with open(updated_file_path, 'w',newline= '\r\n') as file:
                    file.writelines(_data)
                
                # Read the updated file content
                with open(updated_file_path, 'rb') as file:
                    file_content = file.read()

                return file_content, updated_file_name  # Return the file content and name
        except Exception as e:
            st.error(f"An error occurred while updating INP file: {e}")
            return None, None

def main(uploaded_file, modified_values, idx):
    file_content, updated_file_name = update_inp_file(uploaded_file, modified_values, idx)
    if file_content and updated_file_name:
        st.info("INP Updated Successfully!")
        # Provide download link for the updated INP file
        st.download_button(
            label="Download Updated INP",
            data=file_content,
            file_name=updated_file_name,
            mime='text/plain'
        )

if __name__ == "__main__":
    uploaded_file = st.file_uploader("Upload your INP file", type=["inp"])
    main(uploaded_file, modified_values, idx)

'''
-a,b,c,d,e,f,g
-inp

if a is not None or b is not None or c is not None ... g is not None:
    modified_inp = _function(inp, modified_values)
else:
    not modify, just return that content to download furthur

'''