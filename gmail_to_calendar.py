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
import datetime
import re



CALENDAR_EVENT_SUBJECT = "!!일정!!"

def create_calendar_service() :
    SCOPES = ['https://www.googleapis.com/auth/calendar']

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


def parse_time(str) :
    if str is None :
        return None

    year = re.findall('\s*(\d{1,4})년',str)
    month = re.findall('\s*(\d{1,2})월', str)
    day = re.findall('\s*(\d{1,2})일', str)
    pm = re.findall('오후', str)
    hour = re.findall('\s*(\d{1,2})시', str)
    min = re.findall('\s*(\d{1,2})분', str)

    if len(day) == 0 and len(hour) == 0 :
        return None

    today = datetime.date.today()
    year = int(year[0] if len(year) > 0 else today.year)
    month = int(month[0] if len(month) > 0 else today.month)
    day = int(day[0] if len(day) > 0 else today.day)
    hour = int(hour[0] if len(hour) > 0 else 0)
    hour += 12 if len(pm) > 0 else 0
    min = int(min[0] if len(min) > 0 else 0)

    start_date = datetime.datetime(year, month, day, hour, min)
    #print(str, start_date.strftime("%Y-%m-%d %H:%M:%S"))

    return start_date

def create_event(title, start_time_str, duration=None, attendees=None, description=None, location=None):
    start_time = parse_time(start_time_str)

    if duration is None :
        duration = 1

    if start_time is not None :
        end_time = start_time + datetime.timedelta(hours=duration)
    else :
        return None

    if attendees is None :
        attendees = "munhoney@gmail.com"

    timezone = 'Asia/Seoul'  # enter your timezone

    event = {}

    event['summary'] = title
    event['location'] = location
    event['description'] = description
    event['start'] =  {
            'dateTime': start_time.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': timezone,
        }
    event['end'] = {
            'dateTime': end_time.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': timezone,
        }

    if attendees is not None :
        event['attendees'] =  [{'email': attendees}]

    return event


def parse_info_from_gmail(msg) :
    pattern = {"제목": "title",
               "주제": "title",
               "장소": "location",
               "설명": "description",
               "시간": "start_time",
               "일시": "start_time",
               "일정": "start_time",
               "기간": "duration"
               }

    info = dict()

    data = find_data(msg, 'data')
    for d in [data[0]]:
        msg = base64.urlsafe_b64decode(d)
        text = str(msg, 'utf-8')
        text_list = text.split('\n')
        for line in text_list:
            for key, val in pattern.items():
                if key in line and not val in info :
                    regstr = "{}\s*:\s*(.*)".format(key)
                    s = re.findall(regstr, line)
                    if len(s) > 0 :
                        info[val] = s[0].replace("\r","").replace("\n","").strip()

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
            if not CALENDAR_EVENT_SUBJECT in subject[0] :
                continue

            # title, start_time, end_time, location, description = parse_info_from_gmail(msg)
            info = parse_info_from_gmail(msg)
            print(info)

            ## create calendar event based on gmail message
            event = create_event(info.get('title'), info.get('start_time'), info.get('duration'),
                                info.get('attendee'), info.get('description'), info.get('location'))
            if event is None :
                print("ERROR: Event type is wrong", event)
                time.sleep(10)
                continue

            try:
                calendar_service.events().insert(calendarId='primary', body=event, sendNotifications=True).execute()
            except Exception as e:
                print('An error occurred #1', e)
                continue

            ## change label of gmail message (UNREAD -> READ)
            try:
                msg_labels = {'addLabelIds': [], 'removeLabelIds': ['UNREAD']}
                gmail_service.users().messages().modify(userId='me', id=message['id'], body=msg_labels).execute()
            except errors.HttpError as error:
                print('An error occurred #2', error)
                continue

            print("CREATE EVENT: ", info)

        time.sleep(10)

def test_parse_time() :
    test_str = ["2019년 5월 24일",
                "2019년 5월 24일 2시 30분",
                " 2020년 1월3일 2시",
                "2020년 1월 3일 2시 ",
                "2020년 1월 5일 오후 2시",
                "2월 3일 오후 2시",
                "3일 오후 2시",
                "오늘 오후 2시"]
    for str in test_str :
        parse_time(str)

if __name__ == '__main__':
    main()

