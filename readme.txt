Script Overview:
This script will automate eQuest Baseline Automation from the proposed case

Create virtual env
1. pip install virtualenv
2. python -m venv env
3. env\Scripts\activate
4. install required packages
5. to verify 'pip list'
6. pip freeze > requirements.txt



setup env
1. pip install virtualenv
2. python -m venv env
3. env\Scripts\activate  
4. pip install -r requirements.txt

To run code each time
5. To activate env - "env\Scripts\activate"
6. to run code - python main.py   | to run streamlit = streamlit run main.py
7. to deactivate run command  - deactivate

- Steps to convert python file to executable file (exe file) =>
1. Install pyinstaller (If pyinstaller is not installed in the system)
2. execute command- pyinstaller fileName.py --onefile. After this commands- get 2 folders- dist and build. [dist folder has fileName.exe file]
3. execute next command to add logo in executable file- pyinstaller fileName.py --onefile --icon logoName.ico

Note- Initially, logo in the exe file might to seen. In that case, create a copy of that exe file. Then automatically logo is visible.