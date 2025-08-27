import subprocess
import webbrowser
import sys
import os
import tempfile
import shutil

def main():
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)

    # Create temp folder
    temp_dir = tempfile.mkdtemp()

    # Copy individual files
    for filename in ["main.py", "helper.py", "report_ext.py", "script.bat"]:
        src_file = os.path.join(base_path, filename)
        if os.path.exists(src_file):
            dst_file = os.path.join(temp_dir, filename)
            shutil.copy2(src_file, dst_file)

    # Copy src folder
    src_src = os.path.join(base_path, "src")
    src_dest = os.path.join(temp_dir, "src")
    if os.path.exists(src_dest):
        shutil.rmtree(src_dest)
    if os.path.exists(src_src):
        shutil.copytree(src_src, src_dest)

    # Copy database folder
    db_src = os.path.join(base_path, "database")
    db_dest = os.path.join(temp_dir, "database")
    if os.path.exists(db_src):
        if os.path.exists(db_dest):
            shutil.rmtree(db_dest)
        shutil.copytree(db_src, db_dest)

    # Copy doe22 folder
    doe_src = os.path.join(base_path, "doe22")
    doe_dest = os.path.join(temp_dir, "doe22")
    if os.path.exists(doe_src):
        if os.path.exists(doe_dest):
            shutil.rmtree(doe_dest)
        shutil.copytree(doe_src, doe_dest)

    # Add paths so imports work
    sys.path.insert(0, temp_dir)
    sys.path.insert(0, src_dest)

    # Run Streamlit
    script_path = os.path.join(temp_dir, "main.py")
    print("Script exists:", os.path.exists(script_path))

    if os.environ.get("RUNNING_STREAMLIT") != "1":
        os.environ["RUNNING_STREAMLIT"] = "1"
        subprocess.Popen(
            [sys.executable, "-m", "streamlit", "run", script_path, "--server.headless", "true", "--server.port", "8503"],
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        webbrowser.open("http://localhost:8503")
    else:
        import streamlit as st
        st.write("Hello from inside Streamlit app!")



    # Open browser automatically
    webbrowser.open("http://localhost:8503")

if __name__ == "__main__":
    main()
