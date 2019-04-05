from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def create_event_sequence(event):
    events = []
    for hour in range(9, 12):
        for minute in range(59):
            event['summary'] = 'Teste pelo google API ' + str((hour-9)*minute)
            event['start']['dateTime'] = f'2019-04-12T09:{str(minute).zfill(2)}:00'
            event['end']['dateTime'] = f'2019-04-12T09:{str(minute + 1).zfill(2)}:00'

            events.append(service.events().insert(calendarId='caae.dev@gmail.com', body=event).execute())

def delete_next_n_events(n):
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=n, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])
    if not events:
        print('No upcoming events found.')
    else:
        print(f'Deleting the next {len(events)} events')
        for event in events:
            service.events().delete(calendarId='primary', eventId=event['id']).execute()
            # start = event['start'].get('dateTime', event['start'].get('date'))
            # print(start, event['summary'])


def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
    
    event = {
        'summary': 'Teste pelo google API',
        'location': 'Avenida Trabalhador São Carlense, 400',
        'description': 'Dig din dig din dig din.',
        'start': {
            'dateTime': '2019-04-12T09:00:00',
            'timeZone': 'America/Sao_Paulo',
        },
        'end': {
            'dateTime': '2019-04-12T09:01:00',
            'timeZone': 'America/Sao_Paulo',
        },
        'attendees': [
            {'email': 'caae.dev@gmail.com'},
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
            {'method': 'email', 'minutes': 24 * 60},
            {'method': 'popup', 'minutes': 10},
            ],
        },
    }

    # event = service.events().insert(calendarId='caae.dev@gmail.com', body=event).execute()




if __name__ == '__main__':
    main()