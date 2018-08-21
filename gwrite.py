from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

from pprint import pprint
from googleapiclient import discovery

import datetime

#Authorizes read/write permissions
#Returns creds value to authorize reading/writing for other functions
def google_auth():
   
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)

    return creds
    

# Creates a new spreadsheet
# Returns the id of created spreadhsheet
def create_spreadsheet():

    now = datetime.datetime.now()

    creds = google_auth()

    service = discovery.build('sheets', 'v4', credentials=creds)

    title = 'Marketing Dashboard ' + str(now.month) + "/" + str(now.day) + "/" + str(now.year)

    spreadsheet_body = {
    'properties': {'title': title}
    }

    request = service.spreadsheets().create(body=spreadsheet_body)
    response = request.execute()
    
    #Print results to terminal
    pprint(response)

    spreadsheetId = response["spreadsheetId"]

    return spreadsheetId



def create_tab(Id, data):

    creds = google_auth()
    service = discovery.build('sheets', 'v4', credentials=creds)

    spreadsheet_id = Id

    batch_update_spreadsheet_request_body = {
       
        'requests': [data] 

    }

    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_spreadsheet_request_body)
    response = request.execute()

    

def load_tab(Id,tab_name,data):

    creds = google_auth()
    service = discovery.build('sheets', 'v4', credentials=creds)

    service.spreadsheets().values().update(spreadsheetId=Id,
    range=tab_name, body=data, valueInputOption='RAW').execute()

    print('Wrote data to Sheet')

def clear_tab(Id, tab_name):

    creds = google_auth()
    service = discovery.build('sheets', 'v4', credentials=creds)

    
    spreadsheet_id = Id

    range_ = tab_name

    clear_values_request_body = { # Must be empty
        }

    request = service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=range_, body=clear_values_request_body)
    response = request.execute()
