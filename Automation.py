#-------------------------------------------------------------------------------
# Name:             Automation.py
# Purpose:          Automate the points table calculation for the LUGVITC
#
# Organization:     LUGVITC
#
# Created:          21 12 2024
# License:          Apache License Version 2.0
#
# Developed by:     Meit Sant [Github:MT_276]
#-------------------------------------------------------------------------------

print(f'Points table automation for LUGVITC')
print('Developed by     : Meit Sant [Github : MT_276]')
print('Licence          : Apache License Version 2.0')

import openpyxl,sys

# Checking if data file exits 
try:
    with open('data.txt', 'r') as file:
        data = file.readlines()
        path = data[1].split('"')[1]
        print(path)
        
        print("\n[INFO] Using path from data file")
except FileNotFoundError:
    # Taking the path from the user
    print("\n[INFO] Data file not found")
    path = input("[INPUT] Enter the path of the data file > ")
    path = path.replace('\\', '/')
    path = path.replace('"', '')
    
    # Creating a new data file
    with open('data.txt', 'w') as file:
        file.write(f'Path of the file:- \n > "{path}"\n\n')
except:
    print('Error occured while reading data file')
    exit()
    
# Get the file name from the path
print(f"[INFO] Chosen file : {path.split('/')[-1]}")

# Load the workbook
try:
    wb = openpyxl.load_workbook(path)
    print(f"[INFO] Workbook loaded successfully")
except:
    print(f"[ERROR] Error occured while loading the workbook")
    sys.exit()
    
# Load the rubrics sheet
sheet = wb['Rubric']

# Read the leaderboard data into a dictionary
rubrics = {}
for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
    task, points = row
    task = task.split('(')
    task = task[0].strip()
    rubrics[task] = points

# Print the dictionary
print(f"\n[INFO] Rubrics/Points table loaded successfully")




