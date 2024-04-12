import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


class Google:
    service = None

    def authenticate(self):
        """Shows basic usage of the Google Calendar API.
      Prints the start and name of the next 10 events on the user's calendar.
      """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        self.service = build("calendar", "v3", credentials=creds)

    def create_event(self, start_time, stop_time):
        event = {
            'summary': 'Judo Lesgeven',
            'location': 'Klokhofstraat 10, 8400 Oostende',
            'start': {
                'dateTime': start_time.isoformat(),
                'timeZone': 'Europe/Brussels',
            },
            'end': {
                'dateTime': stop_time.isoformat(),
                'timeZone': 'Europe/Brussels',
            },
        }

        event = self.service.events().insert(
            calendarId='a9c321a368136c30ff232df13154fb0737ae7f4c9611009b841f706d69d3bc8b@group.calendar.google.com',
            body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))


if __name__ == '__main__':
    google = Google()
    google.authenticate()

    start_time = datetime.datetime(2024,4, 12, 18, 45)
    stop_time = datetime.datetime(2024,4, 12, 19, 45)
    google.create_event(start_time, stop_time)
