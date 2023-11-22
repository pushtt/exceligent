"""
Project: Automate boring stuffs in excel
- Create the excel file with template
- Fill the excel template with number (by excel functions)
- Format the excel
- Maintain and update the excel
- Move the excel to production
"""

import openpyxl
import pandas as pd
import subprocess
import metric
import job_metadata

"""

Create the workbooks with required sheets to store template, raw and transformation

"""
# Create Deck Sheet
wb = openpyxl.Workbook()
ws_main = wb['Sheet']
ws_main.title = 'Main'

# Create Medata Sheet
wb.create_sheet('_METADATA_')
ws_metdata = wb['_METADATA_']

# Loading the metadata from the script
job_metadata = job_metadata.metadata
for row, (k, v) in enumerate(job_metadata.items(), start=1):
    ws_metdata.cell(row, 1).value = k
    ws_metdata.cell(row, 2).value = v


# set the position of metadata in Main Sheet A2 -> Time, A3 -> Country, A4 -> Job
time_var = ws_main['$A$2']
country_var = ws_main['$A$3']
job_var = ws_main['$A$4']


# Store the value in the position

def find_cell_by_value(sheet, value, col_index):
    for row in range(1, sheet.max_row + 1):
        if sheet.cell(row, col_index).value == value:
            return str(sheet.cell(row, col_index+1).coordinate)
        else:
            continue

time_var.value = ws_metdata[f"{find_cell_by_value(ws_metdata, 'RUN_DATE', 1)}"].value
country_var.value = ws_metdata[f"{find_cell_by_value(ws_metdata, 'FREE_FORM', 1)}"].value
job_var.value = ws_metdata[f"{find_cell_by_value(ws_metdata, 'JOB_ID', 1)}"].value

# Create Raw Sheet
wb.create_sheet('Raw_One')
ws_raw = wb['Raw_One']

"""

Build a template in Main

"""



# Generate A1 as the time array
trailing_week = 5

# Set up the time array
# Column 10 -> Starting the report
# Row 1 -> Time Control
# Row 2 -> Using METADATA date time with Control to generate the time array
# Row 3 -< Convert Row 2 to ISO Week using Excel Formula
# Row 4 -> Convert Row 3 to Year using Excel Formula
# Row 10 -> Starting the report with time array


for index, column in enumerate(range(10, 10+trailing_week+1)):
    time_parameter = ws_main.cell(1, column)
    time_parameter.value = 7*(index-trailing_week)
    ws_main.cell(2, column).value = f'={time_var.coordinate}+{time_parameter.coordinate}'
    ws_main.cell(3, column).value = f'=_xlfn.ISOWEEKNUM({ws_main.cell(2, column).coordinate})'
    ws_main.cell(4, column).value = f'=YEAR({ws_main.cell(2, column).coordinate})'
    ws_main.cell(10, column).value = f'="Wk"&{ws_main.cell(3, column).coordinate}&"-"&RIGHT({ws_main.cell(4, column).coordinate},2)'


# Row 11 -> Set up the metrics
for row, (met, regions) in enumerate(metric.ww_group_by_region.items()):
    print(met)
    for i, region in zip(range(12+row*len(regions), 12 + (row+1)*len(regions) + row), regions):
        print(i, region)
        ws_main.cell(i, 1).value = met
        ws_main.cell(i, 2).value = region

wb.save("main.xlsx")
subprocess.check_call(['open', '-a', 'Microsoft Excel', './main.xlsx'])