'''
Original Author:Nandhana Prakash

Date:19th July 2024

This is to convert an excel file to flatfile(here csv) 

This function is called when any validation is not passed in the given excel file and should be converted to flatfile

The code is not perfect and may contain errors so feel free to provide feedback
'''

#importing pandas
import pandas as pd

def xlsx_to_csv(xlsx_file,csv_file):
    df = pd.read_excel(xlsx_file)
    df.to_csv(csv_file, index=False)

    print(f'Conversion successful. CSV file saved at: {csv_file}')

#Driver code

excel_file_path = "C:\\Users\\Dell\\Documents\\Programming\\Simplain_Project\\voucher_data.xlsx"
csv_file_path = "C:\\Users\\Dell\\Documents\\Programming\\Simplain_Project\\voucher_data.csv"
xlsx_to_csv(excel_file_path,csv_file_path)