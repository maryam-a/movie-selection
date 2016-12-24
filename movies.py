# -*- coding: utf-8 -*-
"""
Created on Thu Dec 22 17:43:03 2016

@author: maryan-a

Referenced the following sources during development:
- https://developers.google.com/sheets/api/quickstart/python
- http://rosettacode.org/wiki/Send_email#Python
- http://naelshiab.com/tutorial-send-email-python/
- https://www.dataquest.io/blog/python-api-tutorial/
- https://www.omdbapi.com/
"""
# Standard libraries
from __future__ import print_function
import os
import random
import getpass
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Google Python Libraries
import httplib2
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# Other External Libraries
import requests

# Constants File
import spreadsheet_id as sid

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

SPREADSHEET_ID = sid.SPREADSHEET_ID
VALUE_INPUT_OPTION = "USER_ENTERED"
EMAIL_SUBJECT = 'Movie Night!'

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

def get_imdb(title, year):
    """ Gets information about the movie.

    Args:
        title (str) - the title of the movie
        year (str) - the year of the movie

    Returns:
        Response, the information about the movie from IMDB
    """
    parameters = {"t": title, "y": year, "plot": "full"}
    response = requests.get("http://www.omdbapi.com/?", params=parameters).json()
    return response

def get_user_information():
    """ Prompts the user for information to send the email

    Returns:
        Name, email, password, to_addr_list, location, time
    """
    pattern = re.compile(r"[\w._\-]+@gmail.com")
    name = input("What is your name? ")

    email = input("What is your gmail address? ")
    while not pattern.match(email):
        email = input("Your email address must end in @gmail.com." +
                      "Please enter a valid gmail address. ")

    password = getpass.getpass(prompt="What is your password? ")
    to_addr_list = [input("Who would you like to send the email to? ")]

    other_email = input("Would you like to send it to another email? If not, type 'no' ")
    while other_email != 'no':
        to_addr_list.append(other_email)
        other_email = input("Would you like to send it to another email? If not, type 'no' ")

    location = input("Where would you like to host the movie? ")
    time = input("What time will the movie start? ")
    return name, email, password, to_addr_list, location, time

def generate_message(name, location, time, imdb_data):
    """ Generates the body of the email to be sent.

    Contains information about the movie, including the year, genre, plot, runtime,
    and list of actors.

    Args:
        name (str) - the name of the user sending the email
        location (str) - the venue of the movie
        time (str) - the time that the movie will be shown
        imdb_data (dict) - the movie data

    Returns:
        Message, the body of the email to be sent
    """
    message = "Hello Friends! \n"
    message += "Tonight, we will be watching " + imdb_data['Title'] + " in the "
    message += location + " at " + time + ". "

    if imdb_data["Response"] == "True":
        message += "Here is some info about the movie:\n"
        message += "Title: " + imdb_data["Title"]
        message += "\nYear: " + imdb_data["Year"]
        message += "\nGenre: " + imdb_data["Genre"]
        message += "\nPlot: " + imdb_data["Plot"]
        message += "\nRuntime: " + imdb_data["Runtime"]
        message += "\nActors: " + imdb_data["Actors"]
    else:
        message += "\n Unfortunately, the movie could not be found in IMDB but "
        message += "you can Google the movie to find out more about it."

    message += "\n\nAll the best,\n" + name
    return message

def send_email(from_addr, to_addr_list, subject, message, password):
    """ Sends an email with a subject and message

    Must be sent using a Gmail email address.

    Args:
        from_addr (str) - the user's Gmail address
        to_addr_list (list(str)) - the emails to which the message will be sent
        subject (str) - the title of the email
        message (str) - the body of the email
        password (str) - the user's password for his/ her Gmail account
    """
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = ", ".join(to_addr_list)
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_addr, password)
    text = msg.as_string()
    server.sendmail(from_addr, ", ".join(to_addr_list), text)
    server.quit()

def main(name, email, password, to_addr_list, location, time):
    """ Selects movie to watch and sends an email notification.
    No email is sent if:
        - the spreadsheet is empty
        - all the movies have been watched
        - the user does not approve of the email to be sent

    Args:
        name (str) - the name of the user sending the email
        email (str) - the user's Gmail address
        password (str) - the user's password for his/ her Gmail account
        to_addr_list (list(str)) - the emails to which the message will be sent
        location (str) - the venue of the movie
        time (str) - the time that the movie will be shown
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discovery_url = ('https://sheets.googleapis.com/$discovery/rest?'
                     'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discovery_url)
    range_name = 'A2:C'
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    values = result.get('values', [])

    if not values:
        print('\nNo data found.')
        return
    else:
        print ('\nRetrieved data from spreadsheet.')
        total = len(values)
        count = 0

        index = random.randint(0, total - 1)
        random_movie = values[index]

        # Has the movie already been watched?
        while len(random_movie) == 3:
            count += 1
            if count >= total:
                print("Sorry. You've watched all the movies.")
                return
            index = random.randint(0, total - 1)
            random_movie = values[index]

        update_range = 'C' + str(index + 2)
        body = {'values': [['x']]}

        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=update_range,
            valueInputOption=VALUE_INPUT_OPTION, body=body).execute()

        imdb_data = get_imdb(random_movie[0], random_movie[1])

        message = generate_message(name, location, time, imdb_data)

        print("\nThis is what will be sent:\n")
        print(message)
        confirmation = input("To confirm, type 'yes'. Otherwise, type 'no' ")

        if confirmation == "yes":
            send_email(email, to_addr_list, EMAIL_SUBJECT, message, password)
            print("\nSent email.")
        else:
            reset_body = {'values': [['']]}
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=update_range,
                valueInputOption=VALUE_INPUT_OPTION, body=reset_body).execute()
            print("\nThe spreadsheet has been reset.")

if __name__ == '__main__':
    name, email, password, to_addr_list, location, time = get_user_information()
    main(name, email, password, to_addr_list, location, time)
