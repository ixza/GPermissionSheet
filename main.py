from google.oauth2 import service_account
from googleapiclient import discovery
from sheetFunctions import sheetFunctions
import config
import pygsheets
import ast



sheetFunctions.fetchData()


with open('names.txt', encoding="utf8") as f:
         emailToNameFile = ast.literal_eval(f.read())
         
sheetsIdList = list(sheetFunctions.protectedSheetsId.values())
userList = list(sheetFunctions.users)
userListRanges = sheetFunctions.getUserRanges(sheetsIdList, userList)
namesList = sheetFunctions.nameCheck(userList, emailToNameFile)



credentials = service_account.Credentials.from_service_account_file(config.CREDENTIALS_PATH)
service = discovery.build('sheets', 'v4', credentials=credentials)
gc = pygsheets.authorize(service_account_file=config.CREDENTIALS_PATH)
wks = gc.open_by_key(config.SPREADSHEET_ID).sheet1


from formatSheets import formatSheets
       
service.spreadsheets().batchUpdate(spreadsheetId=config.SPREADSHEET_ID, body=formatSheets.clearAll).execute() #CLEARS THE FORMAT AND VALUES OF WHOLE SHEET
service.spreadsheets().batchUpdate(spreadsheetId=config.SPREADSHEET_ID, body=formatSheets.enteredSheets).execute()
wks.update_row(7, namesList, col_offset=1)   
wks.update_row(8, userList, col_offset=1)   

offset = 8  
for i, key in enumerate(sheetFunctions.protectedSheets):
    wks.update_col(2, userListRanges[i], row_offset=offset)
    offset+=len(sheetFunctions.protectedSheets[key])+1
    
service.spreadsheets().batchUpdate(spreadsheetId=config.SPREADSHEET_ID, body=formatSheets.adjustColumnWidth).execute() #ADJUST COLUMNS WIDTH