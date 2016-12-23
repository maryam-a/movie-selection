# -*- coding: utf-8 -*-
"""
Created on Thu Dec 22 17:43:03 2016

@author: marya
"""




from __future__ import print_function
import httplib2
import os
import random
import string

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-movies.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'IAP Movies'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-movies.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Shows basic usage of the Sheets API.

    Spreadsheet:
    https://docs.google.com/spreadsheets/d/1lqP2ZNlnYE_V5GehbvW4CgsWqlcdakWE_Q3w2Z5SAVQ/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1lqP2ZNlnYE_V5GehbvW4CgsWqlcdakWE_Q3w2Z5SAVQ'
    rangeName = 'A2:C'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print ('Retrieved data from spreadsheet.')
        
        index, random_movie = random.choice(list(enumerate(values)))
        
        # Has the movie already been watched?
        while len(random_movie) == 3:
            index, random_movie = random.choice(list(enumerate(values)))
            
        updateRange = 'C' + str(index + 2)
        body = { 'values': [['x']] }
        valueInputOption = "USER_ENTERED"
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=updateRange,
            valueInputOption=valueInputOption, body=body).execute()
        
        # get imdb link
        # email should have location, time, movie title, director and imdb link
        

# http://rosettacode.org/wiki/Send_email#Python
# http://naelshiab.com/tutorial-send-email-python/
def sendemail(from_addr, to_addr, subject, message,
              password, smtpserver='smtp.gmail.com:587'):
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
     
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_addr, password)
    text = msg.as_string()
    server.sendmail(from_addr, to_addr, text)
    server.quit()

if __name__ == '__main__':
    main()
