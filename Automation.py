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
#                   Karthik 
#-------------------------------------------------------------------------------

print(f'Points table automation for LUGVITC')
print('Licence          : Apache License Version 2.0')

import openpyxl, sys

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
# Define member dictionary. This dictionary will be saved to another excel file.
member_data = {
    'Reg No': '',
    'Name' : '',
    'Contributions' : '',
    'Events' : '',
    'Points' : 0
}

# Load the rubrics sheet
sheet = wb['Rubric']
# Read the Rubrics data into a dictionary
rubrics = {}
for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
    task, points = row
    task = task.split('(')
    task = task[0].strip()
    rubrics[task] = points



# Reading 'Technical Support Cyber-0-Day' sheet
sheet = wb['Technical Support Cyber-0-Day']
# Read the ith row
row_data = []
for cell in sheet[2]:
    row_data.append(cell.value)
    
# Calculating the points 
member_data['Reg No'] = row_data[1] # Adds the registration number to the dictionary
member_data['Name'] = row_data[0] # Adds the name to the dictionary

if row_data[2] == 'Yes' or row_data[3] == 'Yes':
    member_data['Contributions'] += 'Tech Support, '
    member_data['Events'] += 'Cyber-0-Day 3.0, '
if row_data[2] == 'Yes':
    member_data['Points'] += rubrics['Technical Support']
if row_data[3] == 'Yes':
    member_data['Points'] += rubrics['Technical Support']

print(member_data)
# Searching reg no for other occurances in the workbook
for sheet in wb.sheetnames:
    print(sheet)
    if sheet == 'Technical Support Cyber-0-Day':
        continue
    sheet = wb[sheet]
    for row in sheet.iter_rows(min_row=2, values_only=True):
        print(row)

# Close the workbook
wb.close()

# Saving the data to a new excel file
scraped_values = openpyxl.Workbook()
sheet_of_points = scraped_values.active
sheet_of_points.title = r"Sheet of Points"

headers = list(member_data.keys())
for i in range(len(member_data)):
    temp_cell = sheet_of_points.cell(row = 1, column = i+1)
    temp_cell.value = headers[i]


for i in range(len(member_data)):
    value_cell = sheet_of_points.cell(row = 2, column = i+1 )
    temp_cell = sheet_of_points.cell(row = 1, column = i+1)
    value_cell.value = member_data[temp_cell.value]

scraped_values.save(r"demo.xlsx")







