import socket
socket.setdefaulttimeout(4000) 

import requests

import datetime
import pickle
import os
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import oauth2client
import httplib2

cd = '/home/meltonzheng/Scripts/Excel_to_Google_Sheets/'

gsheet_ids = ['1wJoujb6tDSrR4BX0kUgtyG88lPWR1n4mz7l9M-AtC2Y','1peb2hhOvNb8rmnDN85oCLeppBhsprqK_is7qC-IcoZQ','18z2Zn8UjJujKrWLJMCow9c7imVbCHAyGBzowIfeX56k']
CLIENT_SECRET_FILE = '/home/meltonzheng/Scripts/Excel_to_Google_Sheets/client_secret.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
credentials = flow.run_local_server()
with open(cd + r'refresh.token', 'w+') as f:
    f.write(credentials._refresh_token)
print('Refresh Token:', credentials._refresh_token)
print('Saved Refresh Token to file: refresh.token')
