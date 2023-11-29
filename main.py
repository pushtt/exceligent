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
import report_attributes

"""

Create the workbooks with required sheets to store template, raw and transformation

"""
# Create Deck Sheets
wb = openpyxl.Workbook()
ws_main = wb['Sheet']
ws_main.title = 'Main'

wb.create_sheet('_METADATA_')
ws_metdata = wb['_METADATA_']

wb.create_sheet('Raw_One')
ws_raw = wb['Raw_One']

wb.save('main.xlsx')

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

prefix_gms_metrics = []
for i in report_attributes.time_prefix:
    for metrics in report_attributes.gms_perf_metrics:
        metric_w_prefix = (str(i) + '_' + str(metrics))
        prefix_gms_metrics.append(metric_w_prefix)

ws_raw_columns = report_attributes.date_dimension + report_attributes.region_dimension + report_attributes.seller_origin_dimension + prefix_gms_metrics

for idx, c in enumerate(ws_raw_columns, start=1):
    ws_raw.cell(1, idx).value = c


openpyxl.utils.get_column_letter(ws_raw.max_column)


sheet_title_quote = openpyxl.utils.quote_sheetname(ws_raw.title)
cell_range_list = []
for col in range(1, ws_raw.max_column+1):
    col_letter = openpyxl.utils.get_column_letter(col)
    anchor_cell_range = openpyxl.utils.absolute_coordinate(f'{col_letter}1:{col_letter}{ws_raw.max_row}')
    cell_range_list.append(f'{sheet_title_quote}!{anchor_cell_range}')


for idx, cell_range in zip(range(1, ws_raw.max_column+1), cell_range_list):
    name_defined = ws_raw.cell(1, idx).value
    defn = openpyxl.workbook.defined_name.DefinedName(name_defined, attr_text=cell_range)
    wb.defined_names[name_defined] = defn
"""

Build a template in Main

"""

# Generate A1 as the time array
trailing_week = 5


# Row 11 -> Set up the metrics
for row, (met, regions) in enumerate(metric.ww_group_by_region.items()):
    for i, region in zip(range(12+row*len(regions), 12 + (row+1)*len(regions) + row), regions):
        ws_main.cell(i, 1).value = 'wtd_' + met
        ws_main.cell(i, 2).value = region

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


for idx, row in enumerate(ws_main.iter_rows(min_col=1, max_col=3, min_row=12),start=12):
    print(idx, row)
    for col in range(10, 10+trailing_week+1):
        ws_main.cell(idx, col).value = f'=SUMIFS({row[0].value},activity_year,{ws_main.cell(4, col).coordinate},activity_week,{ws_main.cell(3, col).coordinate},region,"{row[1].value}")'

wb.save("main.xlsx")
subprocess.check_call(['open', '-a', 'Microsoft Excel', './main.xlsx'])
