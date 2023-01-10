from tkinter import UNITS
import gspread

SHEETNAME = 'Client Forms Spreadsheet'
SHEETLINK = 'https://docs.google.com/spreadsheets/d/1EiDgWoVjuS2a7VF7sSjyCTGe48P8ZI_pPsVhMwGKhns/edit'

sheet = gspread.oauth(credentials_filename = r"E:\Coding\formAutomation\client_secrets.json", authorized_user_filename=r"E:\Coding\formAutomation\getSheet.json")

sheet = sheet.open('Client Forms Spreadsheet')

sheet = sheet.get_worksheet(0)

vals = sheet.get_all_values(value_render_option='FORMULA')

def getForm(link : str):
    return link[link.find("d/") + 2:link.find("/viewform")]

def getSheet(link : str):
    return link[link.find('#gid=')+5:link.find('",')]

for x in vals[1:]:
    name = x[0]
    sessionFormLink = getForm(x[2])
    sessionFormSheet = getSheet(x[4])
    collaborativeFormLink = getForm(x[6])
    collaborativeFormSheet = getSheet(x[8])
