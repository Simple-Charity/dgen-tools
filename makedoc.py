
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from authenticator import authenticate

def main(creds, spreadsheet_id, range_name, header_row_index, template_presentation_id, new_presentation_title):
    try:
        # set up the three services and assign them variables
        sheets_service = build("sheets", "v4", credentials=creds)
        slides_service = build("slides", "v1", credentials = creds)
        drive_service = build("drive", "v3", credentials=creds)

        # Extract the spreadsheet values into a variable called 'data'
        data = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()['values']

        # duplicate the template presentation and create relevant variables
        body = {"name": new_presentation_title}
        drive_response = (
            drive_service.files().copy(fileId=template_presentation_id, body=body).execute()
        )
        new_presentation_id = drive_response.get("id")
        new_presentation = slides_service.presentations().get(presentationId = new_presentation_id).execute()
        template_slides = [slide for slide in new_presentation.get("slides")]
        template_ids = [slide['objectId'] for slide in template_slides]


        # Loop through each line of the sheet data below the headers. Add a request to duplicate each template slide, and a request to do a find and replace on every variable name within those slides.

        copy_ids = []

        for i in range(header_row_index+1,len(data)):

            requests = []

            for slide in template_slides:
                requests = requests + [{
                    "duplicateObject": {
                        'objectId': slide["objectId"],
                    }
                }]

            body = {
                "requests": requests
            }

            response = (
                slides_service.presentations().batchUpdate(presentationId = new_presentation_id, body=body).execute()
            )


            # Prepare an array which contains a list of lists of slides, grouped by the row of the spreadsheet to which they correspond. 
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

            body = {
                "requests": requests
            }

            response = (
                slides_service.presentations().batchUpdate(presentationId = new_presentation_id, body=body).execute()
            )

        # Construct requests to put the slides in the right order
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

        body = {
            "requests": requests
        }

        response = (
            slides_service.presentations().batchUpdate(presentationId = new_presentation_id, body=body).execute()
        )

        return "Completed without error."

    except HttpError as err:
        return err
    

# Set the allowed OAuth2.0 scopes that the dgen-tools application can access
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]

credentials = authenticate(scopes)

functional_response = main(
    creds=credentials,
    spreadsheet_id="1gZ0dxwA1ZTvEZCaeUGei0OxQlYxmaE5kYCYvAC4H7Mk",
    range_name="A1:E6",
    header_row_index = 1,
    template_presentation_id="15dC5GTG99YvtVa-ah-Y8Zoi4PhFariETJNNAYuIm-00",
    new_presentation_title="New Test Presentation"
)

print(functional_response)