import os

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText
import base64
import pickle
import time
try:
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
except ImportError, ie:
    print "ERROR: google.auth not found. Install the google.auth modules by typing "
    print "pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib " \
          "into the command line"

# Obtain Authentication or Credentials to access Google Sheets

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/googleapis.com-python-quickstart.json
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.compose']
SUPPORT_SCOPE = ['https://www.googleapis.com/auth/gmail.readonly']
CLIENT_SECRET_FILE = 'Run Files\\client_secret.json'
APPLICATION_NAME = 'Google API for Stats'


def _get_credentials(support=False):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    home_dir = os.path.expanduser('~')
    if support:
        credential = '/support_token.pikle'
        scope = SUPPORT_SCOPE
    else:
        credential = 'token.pickle'
        scope = SCOPES

    credential_dir = os.path.join(home_dir, '.credentials')
    credential_path = credential_dir + credential
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    if os.path.exists(credential_path):
        with open(credential_path, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if support:
            raw_input("You will now be asked to authenticate access for the Support Inbox. "
                      "Please log in as support when prompted. Press enter to continue.")
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, scope)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(credential_path, 'wb') as token:
            pickle.dump(creds, token)

    return creds


MAIL_API = build('gmail', 'v1', credentials=_get_credentials())
SHEETS_API = build('sheets', 'v4', credentials=_get_credentials())
SUPPORT_MAIL_API = build('gmail', 'v1', credentials=_get_credentials(True))


def get_range(rng, sheet_id, sheet_api, dimension='ROWS'):
    return sheet_api.spreadsheets().values().get(spreadsheetId=sheet_id, range=rng,
                                                 majorDimension=dimension).execute().get('values', [])


def _create_message(sender, to, subject, text):
    message = MIMEText(text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_string())}


def _create_draft(service, user_id, message_body):
    try:
        message = {'message': message_body}
        draft = service.users().drafts().create(userId=user_id, body=message).execute()
        return draft
    except HttpError as error:
        print 'An error occurred. Unable to write draft: %s' % error
        return None


def _send_draft(service, user_id, draft):
    service.users().drafts().send(userId=user_id, body={'id': draft['id']}).execute()


def send_message(service, sender, to, subject, text):
    message = _create_message(sender, to, subject, text)
    draft = _create_draft(service, sender, message)  # TODO Does sender work if not "me"?
    _send_draft(service, sender, draft)


def _add_column_value(column, value, value_type):
    # The call addStatValueToColumn(int i, str valueType) is used to create a Sheets API compatible request
    # for each stat. This method essentially builds a column.
    new = {
            "values": [{
                     "userEnteredValue": {value_type: value}
                     }]
            }

    column.append(new)

    return column


# Creates a list of requests for the sheets API. Each entry in statsColumn in a row in the google sheet
# each value in "values" is a cell in the row.
def _create_column(lst):
    column = []

    column = _add_column_value(column, time.strftime("%m/%d/%Y"), "stringValue")

    for item in lst:
        if type(item) is int:
            value_type = "numberValue"
        else:
            value_type = "stringValue"

        _add_column_value(column, item, value_type)

    return column


def _new_column_request(column, col_start, col_end, row_start, row_end):
    # TODO this uses a bathupdate request to execute one request only.
    #  update this to create a single request. then execute all requests with
    #  batch update in case one fails.
    # TODO Determine if row indices are required or change to be more generic new range?
    return {
        "requests": [
            {
                "insertDimension": {
                    "inheritFromBefore": False,
                    "range": {
                        "dimension": "COLUMNS",
                        "startIndex": col_start,
                        "endIndex": col_end,
                        "sheetId": 0
                    }
                }
            },
            {
                "updateCells": {
                    "range": {
                        "startRowIndex": row_start,
                        "endRowIndex": row_end,
                        "sheetId": 0,
                        "startColumnIndex": col_start,
                        "endColumnIndex": col_end
                    },
                    "rows": column,
                    "fields": "*"
                }
            }
        ]
    }


def add_column(lst, service, sheet_id, col_start, col_end, row_start, row_end):
    column = _create_column(lst)
    request = _new_column_request(column, col_start, col_end, row_start, row_end)
    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=request).execute()


def _create_row(values):
    return {
        "values": values,
        "majorDimension": "ROWS"
    }


def update_range(service, sheet_id, rng, values,
                 value_input='USER_ENTERED', value_render='FORMATTED_VALUE'):
    request_body = _create_row(values)
    try:
        service.spreadsheets().values().update(spreadsheetId=sheet_id, body=request_body,
                                               range=rng, valueInputOption=value_input,
                                               responseValueRenderOption=value_render).execute()
    except HttpError, e:
        print "Unable to write Member Specific Stats information"
        raise e  # todo log detailed error


def _create_sort_request(sheet, start_row, end_row, start_col, end_col, order='ASCENDING'):
    return {
        "requests": [{
            "sortRange": {
                "range": {
                    "sheetId": sheet,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_col,
                    "endColumnIndex": end_col
                },
                "sortSpecs": [
                    {
                        "dimensionIndex": 0,  # TODO What is this?
                        "sortOrder": order
                    }
                ]
            }
        }]
    }


def sort_range(service, sheet_id, sheet, start_row, end_row, start_col, end_col, order='ASCENDING'):
    # TODO Make these methods more generic. method to concat request for batch update,
    #  method to execute a batch update etc.
    request = _create_sort_request(sheet, start_row, end_row, start_col, end_col, order)
    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=request).execute()


def get_message(service, user_id, msg_id):
    try:
        response = service.users().messages().get(userId=user_id, id=msg_id, format='metadata').execute()
        message = {'X-GM-THRID': response['threadId'], 'X-Gmail-Labels': response['labelIds']}

        to_find = ['To', 'From', 'Subject', 'Date']

        while len(to_find) > 0:
            found = next(header for header in response['payload']['headers'] if header['name'] in to_find)
            message[found['name']] = found['value']
            to_find.remove(found['name'])

        return message

    except HttpError, e: # todo add stop iterator exeption catch
        pass


def get_messages(service, user_id, search_params=''):
    response = service.users().messages().list(userId=user_id, q=search_params).execute()
    messages = []
    if 'messages' in response:
        messages.extend(response['messages'])

    result = []
    for message in messages:
        result.append(get_message(service, user_id, message['id']))

    return result

    # map(lambda m: )

