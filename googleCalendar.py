import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]
TOKEN_PATH = "credentials/token.json"
CREDENTIALS_PATH = "credentials/credentials.json"

class Google:
    service = None

    def authenticate(self):
        """Authenticates the user and provides credentials for the Google Calendar API."""
        creds = None

        # Check if the token file exists to load credentials
        if os.path.exists(TOKEN_PATH):
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

        # If credentials are not valid or absent, refresh or re-authenticate
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())  # Try refreshing the access token
                except:
                    print("Token refresh failed. Performing full re-authentication.")
                    creds = None  # Invalidate the credentials and force re-authentication
            if not creds:  # If credentials are still not valid, re-authenticate
                flow = InstalledAppFlow.from_client_secrets_file(
                    CREDENTIALS_PATH, SCOPES
                )
                creds = flow.run_local_server(port=0)

            # Save new credentials to the token file for future runs
            with open(TOKEN_PATH, "w") as token:
                token.write(creds.to_json())

        # Build the Google Calendar API service with the valid credentials
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
