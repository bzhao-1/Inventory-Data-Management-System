## July 2022 
## Author: Ben Zhao @Frontage Labs 2022
## Python Script to Automate Inventory Tracking for Biological Instruments at 700 HQ 
## Update Inventory2022 sheet with date and number of items stocked in shelf in lab room, then run script and will return updated spreadsheet with remaining number of items in main storage room 
## When new shipments arrive, add shipment amount to total in "Last" column for row with the item in Inventory2022 sheet then run script will get updated remaining amounts
## Make sure to close updated spreadsheet beforing running script or updates can not be generated 
## Contact @benzhao90@gmail.com or 3474717778 with any questions 

import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference
import numpy as np
import string

dict_df = pd.read_excel('Inventory2022.xlsx', sheet_name=['Gloves.','Cleaning','CENTRIFUGE TUBES','Plate Inventory','Protein Lobind',
'COMBITIPS','RELOAD TIPS','Nova','Cryo Boxes','Cryo Boxes','VWR Solvents','MISC','Boro Glass Fisher'])

def remaining(a, b):
    return b - a

def highlight(val):
    if val <= 20:
        color = 'blue'
    elif val <= 10:
        color = 'yellow'
    elif val <= 5:
        color = 'red'
    else:
        color = 'transparent'
    return 'background-color: {}'.format(color)

def main(dict_df):
    writer = pd.ExcelWriter('Inventory2022Updated.xlsx', engine='xlsxwriter', date_format='dd mmm yyyy')
    for i in dict_df:
        df_gloves = dict_df.get(i)
        df_gloves['Last'] = df_gloves['Last'].fillna(0)
        df_gloves['Remaining in Stock'] = df_gloves.apply(
            lambda x: remaining(x['Boxes put on shelf'], x['Last']), axis=1)
        for index, elem in enumerate(df_gloves['Last']):
            if elem == 0.0:
                df_gloves.at[index,'Last']=df_gloves.at[index-1,'Remaining in Stock']
            df_gloves['Remaining in Stock'] = df_gloves.apply(
                lambda x: remaining(x['Boxes put on shelf'], x['Last']), axis=1)
        df_gloves = df_gloves.drop(['Last'], axis=1, inplace=True)
        df = dict_df.get(i)
        df.style.applymap(highlight, subset=['Remaining in Stock']).to_excel(writer, sheet_name = i)
    
    workbook = writer.book
    fmt_number = workbook.add_format({"num_format": "0"})
    fmt_header = workbook.add_format({'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'font_color': '#FF0000',
    'border': 1})
    for i in dict_df:
        worksheet = writer.sheets[i]
        worksheet.set_column("B:L", 20)
        worksheet.set_column("G:G",20, fmt_number)
    writer.save()

if __name__ == "__main__":
    main(dict_df)