from Google import Create_Service

import pickle as pk
import os 

cd = '/home/meltonzheng/Scripts/Excel_to_Google_Sheets/'

# print(os.path.getsize(cd+'3.pkl'))

with open(cd+'1.pkl', 'rb') as pickle_file:
    rngData1 = pk.load(pickle_file)

with open(cd+'2.pkl', 'rb') as pickle_file:
    rngData2 = pk.load(pickle_file)

with open(cd+'3.pkl', 'rb') as pickle_file:
    rngData3 = pk.load(pickle_file)


# Google Sheet Id
gsheet_ids = ['1wJoujb6tDSrR4BX0kUgtyG88lPWR1n4mz7l9M-AtC2Y','1peb2hhOvNb8rmnDN85oCLeppBhsprqK_is7qC-IcoZQ','18z2Zn8UjJujKrWLJMCow9c7imVbCHAyGBzowIfeX56k']
CLIENT_SECRET_FILE = '/home/meltonzheng/Scripts/Excel_to_Google_Sheets/client_secret.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

worksheet_name_1 = 'MCEI Tableau Data'
worksheet_name_2 = 'Source for Afya Earnings Dashboard'
worksheet_name_3 = 'MCEI and CDCR eConsult Follow-Up MASTER'

service.spreadsheets().values().clear(
    spreadsheetId=gsheet_ids[0],
    range= worksheet_name_1
).execute()

response = service.spreadsheets().values().update(
    spreadsheetId=gsheet_ids[0],
    valueInputOption='RAW',
    range=worksheet_name_1+'!A1',
    body=dict(
        majorDimension='ROWS',
        values=rngData1
    )
).execute()

service.spreadsheets().values().clear(
    spreadsheetId=gsheet_ids[1],
    range=worksheet_name_2
).execute()

response = service.spreadsheets().values().update(
    spreadsheetId=gsheet_ids[1],
    valueInputOption='RAW',
    range=worksheet_name_2+'!A1',
    body=dict(
        majorDimension='ROWS',
        values=rngData2
    )
).execute()

service.spreadsheets().values().clear(
    spreadsheetId=gsheet_ids[2],
    range=worksheet_name_3
).execute()

response = service.spreadsheets().values().update(
    spreadsheetId=gsheet_ids[2],
    valueInputOption='RAW',
    range=worksheet_name_3+'!A1',
    body=dict(
        majorDimension='ROWS',
        values=rngData3
    )
).execute()
