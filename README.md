to run this code first you must change the environment path to your local file

SETUP PROJECT FOLDER
create a folder name "OFAC CODE" \put all the python code in the folder
create one "input" folder
put the header folder in the project folder(you can get the header folder here K:\Reporting\OFAC\OFAC
put the company password csv file inside the project folder
so the the folder will have
OFAC CODE
1. ofac_gui.py
2. ofac_scanner_core.py
3. header_extractor_core.py
4. requirements.txt
5. Company Passwords.csv
6. header
7. input folder

after settle setup the project folder open the IDE of your choosing (Spyder/VSCODE)
go to file(top left) --> open folder --> choose the folder project --> select folder


HOW TO RUN (can be use even if python not installed directly in your local)
1. use vscode(easier to setup)
2. install python extension (ctrl+shift+x then search for python)
3. after installing the python extension the python logo will appear in your vscode(left side)
4. on the top find 'terminal' -> 'new terminal' (make sure the terminal rout to your file directory)
5. then ctrl+shift+p -> create environment -> quick create
6. a .venv folder will appear in your project file
7. then go to terminal paste this -->  .venv\Scripts\activate
8. the filepath will have (.venv) at the front -- if everything okay you are doing well so far
9. copy and paste this in the terminal (you will see the library updating and download the requirement) -->    pip install -r requirements.txt
11. if everything is okay try to run it --> py ofac_gui.py
12. if not okay, idk la TT. jk just reach out to Jia Ming :)
    

