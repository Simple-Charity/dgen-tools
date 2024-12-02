
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from authenticator import authenticate

def linkify(requests, slides, search_var, link_value):
    for slide in slides:
        for element in slide["pageElements"]:
            if ('shape' in element) and ('text' in element['shape']):
                for text_element in element['shape']['text']['textElements']:
                    if 'textRun' in text_element:
                        if search_var in text_element['textRun']['content']:

                            
                            run_start = text_element['startIndex']
                            run_end = text_element['endIndex']

                            if run_start == 97:
                                print(text_element['textRun'])
                                print(run_start)
                                print(run_end)

                            try:
                                requests = requests + [{
                                    "updateTextStyle": {
                                        "objectId": element['objectId'],
                                        "textRange": {
                                            "type": "FIXED_RANGE",
                                            "startIndex": run_start,
                                            "endIndex": run_end,
                                        },
                                        "style": {
                                            "link": {"url": link_value}
                                        },
                                        "fields": "link",
                                    }
                                }]
                            except HttpError as err:
                                print(err)
    return requests

def add_image(requests, slides, search_var, link_value):
    pt52 = {"magnitude": 72, "unit": "PT"}
    for slide in slides:
        for element in slide["pageElements"]:
            if ('shape' in element) and ('text' in element['shape']):
                for text_element in element['shape']['text']['textElements']:
                    if 'textRun' in text_element:
                        if search_var in text_element['textRun']['content']:
                            
                            requests = requests + [{
                               "createImage": {
                                    # "objectId": 'LogoImage',
                                    "url": link_value,
                                    "elementProperties": {
                                        "pageObjectId": slide['objectId'],
                                        "size": {"height": pt52, "width": pt52},
                                        "transform": {
                                            "scaleX": 1,
                                            "scaleY": 1,
                                            "translateX": 504,
                                            "translateY": 10,
                                            "unit": "PT",
                                        },
                                    },
                                }
                            }]

    return requests

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
            drive_service.files().copy(fileId=template_presentation_id, body=body, supportsAllDrives=True).execute()
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

            active_slides = []
            updated_presentation = slides_service.presentations().get(presentationId = new_presentation_id).execute()

            for q in range(len(updated_presentation['slides'])):
                if (q - 1) % (i - header_row_index + 1) == 0:
                    active_slides = active_slides + [updated_presentation['slides'][q]]


            # Prepare an array which contains a list of lists of slides, grouped by the row of the spreadsheet to which they correspond. 
            active_slide_ids = [slide['objectId'] for slide in active_slides]
            copy_ids = copy_ids + [active_slide_ids]

            
            requests = []

            for j in range(len(data[i])):
                field_name = data[header_row_index][j]
                field_value = data[i][j]


                if field_name == "$CEO_Name" and (not data[i][j+1] == ""):
                    requests = linkify(requests, active_slides, "$CEO_Name", data[i][j+1])
                elif field_name == "$COO_Name" and (not data[i][j+1] == ""):
                    requests = linkify(requests, active_slides, "$COO_Name", data[i][j+1])
                elif field_name == "$Director_of_Development_Name" and (not data[i][j+1] == ""):
                    requests = linkify(requests, active_slides, "$Director_of_Development_Name", data[i][j+1])
                elif field_name == "$Board_Chair_Name" and (not data[i][j+1] == ""):
                    requests = linkify(requests, active_slides, "$Board_Chair_Name", data[i][j+1])
                elif field_name == "$Board_Vice_Chair_Name" and (not data[i][j+1] == ""):
                    requests = linkify(requests, active_slides, "$Board_Vice_Chair_Name", data[i][j+1])
                elif field_name == "$Board_Secretary_Name" and (not data[i][j+1] == ""):
                    requests = linkify(requests, active_slides, "$Board_Secretary_Name", data[i][j+1])

                

                if field_name == "$Video_of_Client_Story(ies)":
                    requests = linkify(requests, active_slides, "Video of Client", field_value)
                elif field_name == "$Video_of_Staff":
                    requests = linkify(requests, active_slides, "Video of Staff", field_value)
                elif field_name == "$Impact_Story_Source_Link":
                    requests = linkify(requests, active_slides, "Story Source", field_value)
                elif field_name == "$2021_Annual_Report":
                    requests = linkify(requests, active_slides, "2021 Annual Report", field_value)
                elif field_name == "$2022_Annual_Report":
                    requests = linkify(requests, active_slides, "2022 Annual Report", field_value)
                elif field_name == "$2023_Annual_Report":
                    requests = linkify(requests, active_slides, "2023 Annual Report", field_value)
                elif field_name == "$2021_Audited_Financials":
                    requests = linkify(requests, active_slides, "2021 Financials", field_value)
                elif field_name == "$2022_Audited_Financials":
                    requests = linkify(requests, active_slides, "2022 Financials", field_value)
                elif field_name == "$2023_Audited_Financials":
                    requests = linkify(requests, active_slides, "2023 Financials", field_value)
                elif field_name == "$Strategic_Plan":
                    requests = linkify(requests, active_slides, "Strategic Plan", field_value)
                elif field_name == "$Other_Report_1":
                    requests = linkify(requests, active_slides, "Other Report #1", field_value)
                elif field_name == "$Other_Report_2":
                    requests = linkify(requests, active_slides, "Other Report #2", field_value)
                elif field_name == "$Other_Report_3":
                    requests = linkify(requests, active_slides, "Other Report #3", field_value)



                elif field_name == '$Logo':
                    requests = add_image(requests, active_slides, field_name, field_value)

                    requests = requests + [{
                        "replaceAllText": {
                            'replaceText': "",
                            "pageObjectIds": active_slide_ids,
                            'containsText': {
                                'text': field_name,
                                'matchCase': True,
                            } 
                        }
                    }]
                    

                elif (not field_name == "$CEO_Name") and (not field_name == "$COO_Name") and (not field_name == "$Director_of_Development_Name") and (not field_name == "$Board_Chair_Name") and (not field_name == "$Board_Vice_Chair_Name") and (not field_name == "$Board_Secretary_Name"):
                    requests = requests + [{
                        "replaceAllText": {
                            'replaceText': field_value,
                            "pageObjectIds": active_slide_ids,
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



            requests = []

            for j in range(len(data[i])):

                field_name = data[header_row_index][j]
                field_value = data[i][j]

                if field_name == "$CEO_Name" or field_name == "$COO_Name" or field_name == "$Director_of_Development_Name" or field_name == "$Board_Chair_Name" or field_name == "$Board_Vice_Chair_Name" or field_name == "$Board_Secretary_Name":
                    requests = requests + [{
                        "replaceAllText": {
                            'replaceText': field_value,
                            "pageObjectIds": active_slide_ids,
                            'containsText': {
                                'text': field_name,
                                'matchCase': True,
                            } 
                        }
                    }]

            if not requests == []:

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
    spreadsheet_id="147n6kGgEQ209gJQgfnLUKAgcc6nkWRcNDGSsXatjI0o",
    range_name="A1:IT5",
    header_row_index = 2,
    template_presentation_id="1qoJJ07tqvsY5psYW_kVtik-G69yMOx8h1SBpOdJAatk",
    new_presentation_title="Deliverable Test Run"
)

print(functional_response)