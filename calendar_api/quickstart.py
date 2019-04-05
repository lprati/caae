from __future__ import print_function

import datetime
import pickle
import os.path
import csv
from collections import OrderedDict

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request



# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def create_event(entry_line):
    
    
    event = {
        'summary': '',
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
    
    print(entry_line)
    months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

    date = entry_line['Data da Entrevista']
    hour = entry_line['Horários de Início']
    nome = entry_line['Nome Entrevistador(a)']

    # inicio do evento
    date_values = date.split('/')
    day = date_values[0]
    month = months[int(date_values[1])+1]
    year = date_values[2] 

    start_hour = hour.split(':')[0]
    start_minute = hour.split(':')[1]
    length = 30

    start = datetime.datetime(int(date_values[2]), int(date_values[1]), int(date_values[0]), int(start_hour), int(start_minute), 0)
    end = start + datetime.timedelta(minutes=int(length))

    end_hour = str(end.hour).zfill(2)
    end_minute = str(end.minute).zfill(2)

    event['summary'] = f'{day} de {month} de {year} {hour} BRST - {nome}'
    print(event['summary'])

    event['start']['dateTime'] = f'{year}-{date_values[1]}-{day}T{hour}Z'
    print(event['start']['dateTime'])

    event['end']['dateTime'] = f'{year}-{date_values[1]}-{day}T{end_hour}:{end_minute}:00Z'
    print(event['end']['dateTime'])

    return event

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

def get_csv_as_dict_list(csv_filepath):
    horarios = OrderedDict()
    with open(csv_filepath, mode='r') as file_horarios:
        reader = csv.DictReader(file_horarios)
        keys = reader.fieldnames
        for line in reader:
        
            entry = {}
            for i in range(len(keys)):
                entry[keys[i]] = line[keys[i]]

            if entry != {}:
                horarios[reader.line_num - 1] = entry
    
    return horarios 
    # for i in horarios:
    #     if horarios[i]:
    #         create_event(horarios[i])
    #return horarios


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


    entry = get_csv_as_dict_list('/home/prati/Workbench/caae/calendar_api/horarios.csv')
    event = create_event(entry[1])

    return_value = service.events().insert(calendarId='caae019@gmail.com', body=event).execute()
    print(return_value['iCalUID'])



if __name__ == '__main__':
    main()
