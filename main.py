import pandas as pd
import openpyxl
import os
import numpy as np


def process_data(excel_files, folder_path):
    combined_years = pd.DataFrame()
    for file_name in excel_files:
        print(file_name)
        path_to_file = folder_path + "'" + str(file_name)+"'"
        xls = pd.ExcelFile(folder_path + file_name, engine='openpyxl')
        sheet_names = xls.sheet_names
        for sheet_name in sheet_names:
            print('Loading Sheet: ' + sheet_name)
            sheet_df = pd.read_excel(file_name, sheet_name = sheet_name, engine='openpyxl')
            combined_years = combined_years.append(sheet_df, ignore_index = True)

    combined_dataframe = combined_years
    print(len(combined_dataframe))
    hb_west_dataframe = combined_dataframe.loc[combined_dataframe['Settlement Point Name'] == 'HB_WEST']
    hb_west_dataframe['Delivery Date'] = pd.to_datetime(hb_west_dataframe['Delivery Date']).dt.strftime('%m/%Y')
    hb_west_dataframe = hb_west_dataframe[['Delivery Date', 'Settlement Point Price']]
    print(len(hb_west_dataframe))

    grouped = hb_west_dataframe.groupby('Delivery Date').mean()
    top_3 = grouped.sort_values(by = ['Settlement Point Price'], ascending = [False])
    top_3 = top_3[:3]
    return(top_3)


dir1 = os.getcwd()
folder_path = dir1 + '/Data/'
file_names = os.listdir(folder_path)
excel_files = [file for file in file_names if '.xlsx' in file]
os.chdir(folder_path)
top_3 = process_data(excel_files, folder_path)
top_3.to_excel('Output Top 3.xlsx')
