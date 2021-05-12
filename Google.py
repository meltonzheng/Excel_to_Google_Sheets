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

def refreshToken(client_id, client_secret, refresh_token):
        params = {
                "grant_type": "refresh_token",
                "client_id": client_id,
                "client_secret": client_secret,
                "refresh_token": refresh_token
        }

        authorization_url = "https://www.googleapis.com/oauth2/v4/token"

        r = requests.post(authorization_url, data=params)

        if r.ok:
                return r.json()['access_token']
        else:
                return None
                
def Create_Service(client_secret_file, api_name, api_version, *scopes):
    print(client_secret_file, api_name, api_version, scopes, sep='-')
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERSION = api_version
    SCOPES = [scope for scope in scopes[0]]
    print(SCOPES)

    cred = None
    ###########################################
    # Use below for a new refresh token
    #############################################################
    # flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
    # credentials = flow.run_local_server()
    # with open('refresh.token', 'w+') as f:
    #     f.write(credentials._refresh_token)
    # print('Refresh Token:', credentials._refresh_token)
    # print('Saved Refresh Token to file: refresh.token')
    #############################################################
    
    with open(cd+'refresh.token','r') as f:
    	refresh_token = f.read()
    
    cred = oauth2client.client.GoogleCredentials(None,'256022659029-qvu7b42lt8mjdnlm1f7ufic184r9id43.apps.googleusercontent.com','dgFnqjcNaYUobwmgoIeEaraO',refresh_token,None,"https://accounts.google.com/o/oauth2/token",None)
    http = cred.authorize(httplib2.Http())
    cred.refresh(http)
    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
        print(API_SERVICE_NAME, 'service created successfully')
        return service
    except Exception as e:
        print(e)
    return None

def convert_to_RFC_datetime(year=1900, month=1, day=1, hour=0, minute=0):
    dt = datetime.datetime(year, month, day, hour, minute, 0).isoformat() + 'Z'
    return dt
