from tkinter import UNITS
import gspread

DONTMAKEVOUCHER = 0
PRIMARYCLIENT = 1
RENEWALDATE = 2
UNITSPRIORTOTHISSHEET = 3
TOTALUNITSWITHTHISSHEET = 4
CASEMANAGER = 5
CLIENT = 6
PRIMARYCLINICIAN = 7
INVOICENUMBER = 8
TOTALUNITS = 9




sheet = gspread.oauth(credentials_filename = r"E:\Coding\formAutomation\client_secrets.json", authorized_user_filename=r"E:\Coding\formAutomation\token1.json")

sheet = sheet.open('MCSS Submission Spreadsheets 2022')

sheet = sheet.get_worksheet(1)

vals = sheet.get_all_values()

start = 0

for i, line in enumerate(vals):
    if line[0].lower().replace(' ', '') ==  'MCSS - NEED VOUCHERS'.lower().replace(' ', ''):
        start = i
    
    elif line[0].lower().replace(' ', '') == "NO VOUCHERS BELOW THIS LINE".lower().replace(' ', ''):
        break

clientGroup = []

i=0
tmp = []

while i < len(vals[start:i]):
