#################################################################
# Gmail to Calendar Program
# This program can create calendar event based on gmail message
# Author: SungTae Moon (munhoney@gmail.com)
#################################################################

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import time
import email
from apiclient import errors
import datefinder
import datetime


def create_calendar_service() :
    SCOPES = ['https://www.googleapis.com/auth/calendar']

    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token_calendar.pickle'):
        with open('token_calendar.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token_calendar.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service

def create_gmail_service() :
    SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds = None
    if os.path.exists('token_gmail.pickle'):
        with open('token_gmail.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token_gmail.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    return service


def data_encoder(text):
    if len(text)>0:
        message = base64.urlsafe_b64decode(text)
        message = str(message, 'utf-8')
        message = email.message_from_string(message)
    return message

def find_data(dictionary, key_to_find, depth=0, result=None):
    if result is None:
        result = list()

    for key, value in dictionary.items():
        if key_to_find == key:
            result.append(value)
        elif isinstance(value, list) and key != "headers":
            for v in value:
                if isinstance(v, dict):
                    found = find_data(v, key_to_find, depth=depth + 1, result=result)
                    if found is not None:
                        return found
        elif isinstance(value, dict):
            found = find_data(value, key_to_find, depth=depth + 1, result=result)
            if found is not None:
                return found

    if depth == 0:
        return result

def create_event(title, start_time_str, duration=None, attendees=None, description=None, location=None):
    start_time_str = start_time_str.replace(u'\xa0', "").replace(" ", "").replace("년", "-").replace("월", "-").replace("일", " ").replace("시",":").replace("분","")
    matches = list(datefinder.find_dates(start_time_str))
    
    if duration is None :
        duration = 1

    if len(matches):
        start_time = matches[0]
        end_time = start_time + datetime.timedelta(hours=duration)

    if attendees is None :
        attendees = "munhoney@gmail.com"

    timezone = 'Asia/Seoul'  # enter your timezone

    event = {
        'summary': title,
        'location': location,
        'description': description,
        'start': {
            'dateTime': start_time.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': timezone,
        },
        'end': {
            'dateTime': end_time.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': timezone,
        },
        'attendees': [
            {'email': attendees},
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10},
            ],
        },
    }

    return event

def parse_info_from_gmail(msg) :
    pattern = {"제목;": "title",
               "주제;": "title",
               "장소;": "location",
               "설명;": "description",
               "시간;": "start_time",
               "일시;": "start_time",
               "기간;": "duration"
               }

    info = dict()

    data = find_data(msg, 'data')
    for d in [data[0]]:
        msg = base64.urlsafe_b64decode(d)
        text = str(msg, 'utf-8')
        text_list = text.split('\n')
        for t in text_list:

            for key, val in pattern.items():
                if key in t :
                    s = t.split(key)
                    info[val] = s[1].replace("\r","").replace("\n","").strip()

    return info

def main():

    ## process to access the gmail message using token_gmail.pickle
    gmail_service = create_gmail_service()

    ## process to access the google calendar event using token_calendar.pickle
    calendar_service = create_calendar_service()

    while (True) :

        ## read gmail message list
        results = gmail_service.users().messages().list(userId='me', labelIds=['UNREAD', 'INBOX']).execute()
        messages = results.get('messages', [])

        for message in messages:
            ## read gmail message
            msg = gmail_service.users().messages().get(userId='me', id=message['id']).execute()

            headers = msg["payload"]["headers"]
            subject = [i['value'] for i in headers if i["name"] == "Subject"]

            ## check some pattern for calendar event
            if not "!!일정!!" in subject[0] :
                continue

            # title, start_time, end_time, location, description = parse_info_from_gmail(msg)
            info = parse_info_from_gmail(msg)
            print(info)

            ## create calendar event based on gmail message
            event = create_event(info.get('title'), info.get('start_time'), info.get('duration'),
                                info.get('attendee'), info.get('description'), info.get('location'))
            try:
                calendar_service.events().insert(calendarId='primary', body=event, sendNotifications=True).execute()
            except Exception as e:
                print('An error occurred', e)
                continue

            ## change label of gmail message (UNREAD -> READ)
            try:
                msg_labels = {'addLabelIds': [], 'removeLabelIds': ['UNREAD']}
                gmail_service.users().messages().modify(userId='me', id=message['id'], body=msg_labels).execute()
            except errors.HttpError as error:
                print('An error occurred #2', error)
                continue

        time.sleep(10)


if __name__ == '__main__':
    main()