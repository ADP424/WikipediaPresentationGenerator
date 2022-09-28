from __future__ import print_function

import os.path
import random
import string

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials


# If modifying these scopes, delete the file token.json.

SCOPES = ['https://www.googleapis.com/auth/presentations']


def get_credentials():
    """
    99% of this function's code was provided by Google
    Checks user credentials and requests Google login for the user if it's their first use of the program
    Returns slides service
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is created automatically when the
    # authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the slides_credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('slides', 'v1', credentials=creds)


class presentation:
    def __init__(self, service):
        self.service = service

    def create_presentation(self, title: str):
        """
        Creates a blank presentation
        """
        slides_service = self.service
        body = {'title': title}
        pres = slides_service.presentations().create(body=body).execute()
        print('Created presentation with ID: ' + str(pres.get('presentationId')))

        return pres

    def create_slide(self, presentation_id, page_id, title: str, textBody: str):
        """
        Creates a blank slide and then populates it with text boxes containing title and textBody
        List of layouts: https://developers.google.com/apps-script/reference/slides/predefined-layout
        """

        title_id = "".join(random.choice(string.ascii_letters) for _ in range(10))
        body_id = "".join(random.choice(string.ascii_letters) for _ in range(10))

        slides_service = self.service
        requests = [
            {
                'createSlide': {
                    'objectId': page_id,
                    'insertionIndex': '1',
                    'slideLayoutReference': {
                        'predefinedLayout': 'BLANK'
                    }
                }
            },

            # creates the title and modifies its text format
            {
                'createShape': {
                    'objectId': title_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': page_id,
                        'size': {
                            'height': {
                                'magnitude': 40,
                                'unit': 'PT'
                            },
                            'width': {
                                'magnitude': 720,
                                'unit': 'PT'
                            }
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 0,
                            'translateY': 20,
                            'unit': 'PT'
                        }
                    }
                }
            },

            {
                'insertText': {
                    'objectId': title_id,
                    'insertionIndex': 0,
                    'text': title
                }
            },

            {
                'updateTextStyle': {
                    'objectId': title_id,
                    'style': {
                        'fontFamily': 'Times New Roman',
                        'fontSize': {
                            'magnitude': 40,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'fontFamily,fontSize'
                }
            },

            {
                'updateParagraphStyle': {
                    "objectId": title_id,
                    "style": {
                        "alignment": "CENTER"
                    },
                    "fields": 'alignment',
                }
            },


            # creates the text body and modifies its text format
            {
                'createShape': {
                    'objectId': body_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': page_id,
                        'size': {
                            'height': {
                                'magnitude': 500,
                                'unit': 'PT'
                            },
                            'width': {
                                'magnitude': 660,
                                'unit': 'PT'
                            }
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 30,
                            'translateY': 100,
                            'unit': 'PT'
                        }
                    }
                }
            },

            {
                'insertText': {
                    'objectId': body_id,
                    'insertionIndex': 0,
                    'text': textBody
                }
            },

            {
                'updateTextStyle': {
                    'objectId': body_id,
                    'style': {
                        'fontFamily': 'Times New Roman',
                        'fontSize': {
                            'magnitude': 18,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'fontFamily,fontSize'
                }
            },

            # bulletPreset enums found here:
            # https://developers.google.com/slides/api/reference/rest/v1/presentations/request#bulletglyphpreset
            {
                'createParagraphBullets': {
                    'objectId': body_id,
                    'textRange': {
                        'type': 'ALL'
                    },
                    'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                }
            }
        ]

        body = {'requests': requests}
        response = slides_service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
        create_slide_response = response.get('replies')[0].get('createSlide')
        print('Created slide with ID: ' + str(create_slide_response.get('objectId')))

        return create_slide_response

    def set_slide_title(self, presentation_id, title_id, title_text):
        slides_service = self.service
        requests = [{
            'insertText': {
                'objectId': title_id,
                'insertionIndex': 0,
                'text': title_text
            }
        }]

        # Execute the requests.
        body = {'requests': requests}
        response = slides_service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
        print('Replaced text in shape with ID: {0}'.format(title_id))
        return title_id
