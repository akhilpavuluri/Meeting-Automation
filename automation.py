import openpyxl
import requests
from datetime import datetime

# Load data from Excel sheet
wb = openpyxl.load_workbook('meeting_data.xlsx')
sheet = wb.active

# Specify column indexes for different data
platform_column = 1
link_column = 2
username_column = 3
password_column = 4
datetime_column = 5  # Column index for date and time

# Iterate through rows in the Excel sheet
for row in sheet.iter_rows(min_row=2):
    # Extract data from Excel sheet
    platform = row[platform_column - 1].value
    link = row[link_column - 1].value
    username = row[username_column - 1].value
    password = row[password_column - 1].value
    datetime_str = row[datetime_column - 1].value # Date and time from Excel sheet

    # Start the meeting based on the platform
    if platform.lower() == 'zoom':
        zoom_api_key = '<YOUR_ZOOM_API_KEY>'
        zoom_api_secret = '<YOUR_ZOOM_API_SECRET>'

        headers = {
            'Content-Type': 'application/json',
        }
        data = {
            'topic': 'Meeting',
            'type': 1,
            'start_time': datetime.isoformat(),
            'timezone': 'Asia/Kolkata',
            'settings': {
                'join_before_host': True
            }
        }

        response = requests.post(f'https://api.zoom.us/v2/users/me/meetings', json=data, headers=headers, auth=(zoom_api_key, zoom_api_secret))
        if response.status_code == 201:
            join_url = response.json()['join_url']
            print(f"Zoom meeting started. Join URL: {join_url}")
        else:
            print("Failed to start Zoom meeting.")

    elif platform.lower() == 'google meet':
        google_meet_api_key = '<YOUR_GOOGLE_MEET_API_KEY>'
        try:
            datetime = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')  # Convert the string to a datetime object
        except ValueError:
            print("Invalid datetime format in the Excel sheet.")
            continue
        # Convert datetime to string format

        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {google_meet_api_key}'
        }
        from datetime import datetime

        start_datetime = datetime.isoformat()

        data = {
            'conferenceData': {
                'createRequest': {
                    'conferenceSolutionKey': {
                        'type': 'hangoutsMeet'
                    },
                    'requestId': '1234567890'
                }
            },
            'start': {
                'dateTime': start_datetime,
                'timeZone': 'Asia/Kolkata'
            },
            'end': {
                'dateTime': datetime.isoformat(),  # Update this with the end date/time
                'timeZone': 'Asia/Kolkata'
            },
            'summary': 'Meeting'
        }

        response = requests.post('https://www.googleapis.com/calendar/v3/calendars/primary/events', json=data, headers=headers)
        if response.status_code == 200:
            conference_data = response.json().get('conferenceData')
            join_url = conference_data.get('entryPoints')[0].get('uri')
            print(f"Google Meet meeting started. Join URL: {join_url}")
        else:
            print("Failed to start Google Meet meeting.")

    elif platform.lower() == 'microsoft teams':
        teams_access_token = '<YOUR_MICROSOFT_TEAMS_ACCESS_TOKEN>'

        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {teams_access_token}'
        }
        data = {
            'startDateTime': datetime.isoformat(),
            'endDateTime': datetime.isoformat(),  # Update this with the end date/time
            'subject': 'Meeting',
            'participants': {
                'organizer': {
                    'identity': {
                        'user': {
                            'displayName': 'Organizer Name',
                            'tenantId': '<YOUR_MICROSOFT_TEAMS_TENANT_ID>'
                        }
                    }
                }
            }
        }

        response = requests.post('https://graph.microsoft.com/v1.0/me/onlineMeetings', json=data, headers=headers)
        if response.status_code == 201:
            join_url = response.json().get('joinWebUrl')
            print(f"Microsoft Teams meeting started. Join URL: {join_url}")
        else:
            print("Failed to start Microsoft Teams meeting.")

    else:
        print(f"Unknown platform: {platform}")
