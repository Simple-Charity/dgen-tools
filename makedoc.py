
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]

spreadsheet_id = "1gZ0dxwA1ZTvEZCaeUGei0OxQlYxmaE5kYCYvAC4H7Mk"
sheet_id = "2056817935"
range_name = "A1:E6"

presentation_id = "15dC5GTG99YvtVa-ah-Y8Zoi4PhFariETJNNAYuIm-00"
creds = None

# The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", scopes)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)

        creds = flow.run_local_server(port=0)

    # Save the credentials for the next run
    with open("token.json", "w") as token:
        token.write(creds.to_json())


try:
    sheets_service = build("sheets", "v4", credentials=creds)
    slides_service = build("slides", "v1", credentials = creds)

    data = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()['values']


    # 'slides' is an array of the slides in the template presentation
    template_presentation = slides_service.presentations().get(presentationId = presentation_id).execute()

    # duplicate the template presentation and create relevant variables
    drive_service = build("drive", "v3", credentials=creds)
    body = {"name": "New Test Presentation"}
    drive_response = (
        drive_service.files().copy(fileId=presentation_id, body=body).execute()
    )
    new_presentation_id = drive_response.get("id")
    new_presentation = slides_service.presentations().get(presentationId = new_presentation_id).execute()
    template_slides = [slide for slide in new_presentation.get("slides")]

    template_ids = [slide['objectId'] for slide in template_slides]

    copy_ids = []

    for i in range(2,len(data)):

        requests = []

        for slide in template_slides:
            requests = requests + [{
                "duplicateObject": {
                    'objectId': slide["objectId"],
                }
            }]

        body = {"requests": requests}

        response = (
            slides_service.presentations().batchUpdate(presentationId = new_presentation_id, body=body).execute()
        )

        active_slides = [reply['duplicateObject']['objectId'] for reply in response['replies']]
        copy_ids = copy_ids + [active_slides]
        
        requests = []

        for j in range(len(data[i])):
            field_name = data[1][j]
            field_value = data[i][j]


            requests = requests + [{
                "replaceAllText": {
                    'replaceText': field_value,
                    "pageObjectIds": active_slides,
                    'containsText': {
                        'text': field_name,
                        'matchCase': True,

                    } 
                }
            }]

        body = {"requests": requests}

        response = (
            slides_service.presentations().batchUpdate(presentationId = new_presentation_id, body=body).execute()
        )


    requests = []
    count = 0
    for id_set in copy_ids:
        requests = requests + [{
            'updateSlidesPosition': {
                'slideObjectIds': id_set,
                'insertionIndex': count,
            }
        }]

        count = count + len(id_set)

    for template_slide_id in template_ids:
        requests = requests + [{
            'deleteObject': {
                'objectId': template_slide_id,
            }
        }]

    body = {"requests": requests}

    response = (
        slides_service.presentations().batchUpdate(presentationId = new_presentation_id, body=body).execute()
    )





except HttpError as err:
    print(err)
