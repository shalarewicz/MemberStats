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

from util import print_error

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


def get_range(rng, sheet_id, sheet_api, dimension='ROWS', values_only=True):
    """
    Obtains a list of values for the given spreadsheet range
    :param rng: Valid range in A1 notation or a defined named range
    :param sheet_id: Sheet from which then range is obtained.
    :param sheet_api: Sheets API service used
    :param dimension: dimension for representing the data in a list. (default = 'ROWS')
    :param values_only: True if only a list of dimension values should be returned else a ValueRange object is
    returned (default = True)
    :return: a list of either rows or columns depending on the dimension used.
    """
    try:
        if values_only:
            return sheet_api.spreadsheets().values().get(spreadsheetId=sheet_id, range=rng,
                                                         majorDimension=dimension).execute().get('values', [])
        else:
            return sheet_api.spreadsheets().values().get(spreadsheetId=sheet_id, range=rng,
                                                         majorDimension=dimension).execute()
    except HttpError, e:
        print_error('Error: Could not get range: ' + str(rng) + ' from sheet ' + str(sheet_id))
        raise e


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
        print_error('Error: Failed to create. Please see stats_email.txt for draft and send manually.')
        raise error


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
    try:
        service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=request).execute()
    except HttpError, e:
        print_error('Error: add column to sheet: ' + str(sheet_id))
        raise e


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
        print_error('Error: Failed to update range: ' + str(rng) + ' on sheet: ' + str(sheet_id))
        raise e


def update_request(sheet_id, row_data, start_row, end_row, start_col, end_col, ):
    return {
      "updateCells": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": start_row,
            "endRowIndex": end_row,
            "startColumnIndex": start_col,
            "endColumnIndex": end_col
        },
        "rows": row_data,
        "fields": "*"
      }
    }


def sort_request(sheet, sort_index, start_row, end_row, start_col, end_col, order='ASCENDING'):
    return {"sortRange": {
                "range": {
                    "sheetId": sheet,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_col,
                    "endColumnIndex": end_col
                },
                "sortSpecs": [
                    {
                        "dimensionIndex": sort_index,
                        "sortOrder": order
                    }
                ]
            }}


def get_message(service, user_id, msg_id, labels):
    to_find = ['To', 'From', 'Subject', 'Date']
    try:
        response = service.users().messages().get(userId=user_id, id=msg_id, format='metadata').execute()
        message = {'X-GM-THRID': response['threadId'], 'X-Gmail-Labels': response['labelIds']}
        message['X-Gmail-Labels'] = map(lambda l: labels[l], message['X-Gmail-Labels'])

        while len(to_find) > 0:
            found = next(header for header in response['payload']['headers'] if header['name'] in to_find)
            encoded = found['value'].encode('ascii', 'ignore')
            message[found['name']] = encoded
            to_find.remove(found['name'])

        return message
    except StopIteration, se:
        print "Message incomplete: missing " + str(to_find)
    except HttpError, e:
        print_error('Error: Failed to retrieve message: ' + msg_id)
        raise e


def get_labels(service, user_id='me'):
    """
    Obtain a dictionary of label IDs mapped to their respective label name.
    :param service: Mail API service used to obtain the lables. Must have read access to user's mail
    :param user_id: default to 'me'
    :return: dictionairy (key, value) = (id, name)
    """
    try:
        response = service.users().labels().list(userId=user_id).execute()
        labels = {}

        for item in response['labels']:
            labels[item['id']] = item['name']

        return labels
    except HttpError, e:
        print_error('Error: Failed to retrieve labels for user: ' + str(user_id))
        raise e


def get_messages(service, user_id='me', query=''):
    """
    Obtains a list of gmail messages containing Thread ID, Subject, To, From and Labels
    :param service: Mail API service used to obtain the messages. Must have read access to user's mail
    :param user_id: default to 'me'
    :param query:
    :return: list of messages. message.keys() = 'X-GM-THRID' , Subject, To, From and 'X-Gmail-Labels'
    """
    try:
        response = service.users().messages().list(userId=user_id, q=query).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id, q=query,
                                                       pageToken=page_token).execute()
            messages.extend(response['messages'])
        labels = get_labels(service, user_id)
        result = []
        for message in messages:
            new_message = get_message(service, user_id, message['id'], labels)
            if new_message is not None:
                result.append(new_message)
        return result

    except HttpError, e:
        print_error('Error: Failed to retrieve messages for: ' + str(user_id) + 'using query: ' + str(query))
        raise e


def duplicate_sheet_request(sheet_id, new_title, insert_index):
    return {'duplicateSheet': {
        'sourceSheetId': sheet_id,
        'insertSheetIndex': insert_index,
        'newSheetName': new_title
        }
    }


def delete_named_range_request(service, spreadsheet_id, rng):
    try:
        named_ranges = service.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=rng).execute().get(
            'namedRanges', [])

        request_body = []
        for named_range in named_ranges:
            range_id = named_range['namedRangeId']
            request_body.append({"deleteNamedRange": {"namedRangeId": range_id}})

        return request_body
    except HttpError, e:
        print_error('Error: Failed to retrieve named ranges from range: ' + str(rng) +
                    'on sheet: ' + str(spreadsheet_id))
        raise e


def remove_formulas(service, spreadsheet_id, rng):
    values = get_range(rng, spreadsheet_id, service, 'COLUMNS', False)
    try:
        service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, valueInputOption='RAW',
                                               range=values['range'], body=values).execute()
    except HttpError, e:
        print_error('Error: Failed to remove formulas from range: ' + str(rng) + 'on sheet: ' + str(spreadsheet_id))
        raise e


def clear_ranges(service, spreadsheet_id, ranges):
    """

    :param service:
    :param spreadsheet_id:
    :param ranges:
    :return:
    """
    request_body = {'ranges': ranges}
    try:
        service.spreadsheets().values().batchClear(spreadsheetId=spreadsheet_id, body=request_body).execute()
    except HttpError, e:
        print_error('Error: Failed to clear ranges: ' + str(ranges) + 'on sheet: ' + str(spreadsheet_id))
        raise e


def spreadsheet_batch_update(service, spreadsheet_id, requests):
    request_body = {'requests': requests}

    try:
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request_body).execute()
    except HttpError, e:
        request_names = []
        for request in requests:
            request_names.append(request.keys())
        print_error('Error: Batch update failed. Sheet: ' + str(spreadsheet_id) + 'Requests: ' + str(request_names))
        raise e
