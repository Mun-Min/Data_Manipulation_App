import gspread
import os
import streamlit as st 
import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook
from oauth2client.service_account import ServiceAccountCredentials

def getGoogleSheet(url): 

    # define the scope for the Google Sheets API
    scope = ['https://www.googleapis.com/auth/spreadsheets']

    # create a ServiceAccountCredentials object using the private key file
    credentials = ServiceAccountCredentials.from_json_keyfile_name('./symmetric-flare-370800-0ced80d44db6.json', scope)

    # authenticate with Google and open the Sheets file
    gc = gspread.authorize(credentials)
    sh = gc.open_by_url(url)

    # get the first sheet in the file
    worksheet = sh.get_worksheet(0)

    # get the values of the first sheet as a 2D array
    values = worksheet.get_all_values()

    # create a new Excel Workbook
    wb = Workbook()

    # get the active sheet in the new Excel file
    ws = wb.active

    # iterate over the values and write them to the Excel sheet
    for row in values:
        ws.append(row)

    # Get the path to the current user's home directory
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

    # Create the file path for the Excel file
    file_path = os.path.join(downloads_folder, 'file.xlsx')

    # Save the Excel file to the downloads folder
    wb.save(file_path)

    # Create a link to download the Excel file
    st.markdown('Successfully downloaded Excel File to your Downloads Folder...')


def getFilePath(): 
    
    # Get the path to the current user's home directory
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

    # Create the file path for the Excel file
    file_path = os.path.join(downloads_folder, 'file.xlsx')

    return file_path

def viewData():
    df = pd.read_excel(getFilePath())
    st.dataframe(df)