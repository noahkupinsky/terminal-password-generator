from __future__ import print_function

import os.path
import subprocess
import sys, getopt

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# The ID and range of a sample spreadsheet.
id_file = open("spreadsheetID.txt")
SPREADSHEET_ID = id_file.readline()
SHEETS_API = None
GET_MESSAGE = "Password:"
MAKE_MESSAGE = "Successfully created account. Password:"
SEARCH = 0
DICTIONARY = 1
NAME = 0
EMAIL = 1
MAX = 2
DASHLESS = 3
MAX_DEFAULT = 255
DASHLESS_DEFAULT = False


def get_credentials():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def initialize_sheets_api():
    global SHEETS_API
    try:
        service = build('sheets', 'v4', credentials=get_credentials())
        SHEETS_API = service.spreadsheets()
    except HttpError as err:
        print("failed to initialize sheets api")
        print(err)


def get_sheet_values(sheet_name):
    try:
        result = SHEETS_API.values().get(spreadsheetId=SPREADSHEET_ID,
                                         range=sheet_name).execute()
        return result.get('values', [])
    except HttpError as err:
        print("failed to get sheet %s" % sheet_name)
        print(err)


def find_data_from_match(sheet_values, search_column, search_data, return_column):
    for i in range(len(sheet_values)):
        if sheet_values[i][search_column].lower() == search_data.lower():
            return sheet_values[i][return_column]
    return None


def to_password(key, max_characters, dashless):
    # pad the key in case we don't enter in enough characters
    key += "___"
    # readability variables
    # get alphabet sheet values
    alphabet = get_sheet_values("Alphabet")
    # get tag format
    tag_format = find_data_from_match(alphabet, SEARCH,
                                      "tag", DICTIONARY)
    # calculate the number of vowels
    num_vowels = sum([key.count(x) for x in "aeiou"])
    # create the tag
    tag = tag_format.replace("$", str(num_vowels))
    password_array = [tag]
    # fill out remaining password words
    for i in range(3):
        character = key[i]
        # if not alphabetic character we want to replace with 1-3 depending on position
        if character.lower() not in "abcdefghijklmnopqrstuvwxyz":
            character = str(i + 1)
        password_array.append(find_data_from_match(alphabet, SEARCH,
                                                   character, DICTIONARY))
    # join password array into a single string
    joiner = "" if dashless else "-"
    password = joiner.join(password_array)
    # enforce character limit if there is one
    if max_characters < len(password):
        password = password[:max_characters]

    return password


def find_row_starting_with(name, error_when_not_found=True):
    accounts = get_sheet_values("Accounts")
    for i in range(len(accounts)):
        row = accounts[i]
        if row[NAME].lower().startswith(name.lower()):
            row[MAX] = int(row[MAX])
            row[DASHLESS] = (row[DASHLESS].lower == "true") #parse the string in the table
            return row, i
    raise FileNotFoundError("Cannot find Account with that name")


def get_password_from_row_data(row):
    key = row[NAME][:3]
    return to_password(key, row[MAX], row[DASHLESS])


def delete_row(row_index):
    delete_request_body = {
        'requests': [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": 0,
                        "dimension": "ROWS",
                        "startIndex": row_index,
                        "endIndex": row_index + 1
                    }
                }
            },
        ]
    }
    SHEETS_API.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=delete_request_body).execute()
    print("Successfully deleted row")


def process_make_request(name, email, max_characters, dashless, override):
    try:
        _, index = find_row_starting_with(name, False)
        if override:
            delete_row(index)
        else:
            raise Exception("Account name already exists")
    except FileNotFoundError:
        pass# didn't find existing account - not an issue though
    row = [name, email, max_characters, dashless]
    SHEETS_API.values().append(
        spreadsheetId=SPREADSHEET_ID, range="Accounts",
        valueInputOption="USER_ENTERED", body={'values': [row]}).execute()
    return get_password_from_row_data(row)


def print_password(password, message, copy):
    print(message)
    print(password)
    if copy:
        task = subprocess.Popen(
            ['pbcopy'],
            stdin=subprocess.PIPE,
            close_fds=True
        )
        task.communicate(input=password.encode('utf-8'))


def get_options(args):
    opts, args = getopt.getopt(args, "codm:e:", ["copy", "override", "dashless", "maxchars=", "email="])
    email = ""
    max_characters = MAX_DEFAULT
    dashless = DASHLESS_DEFAULT
    override = False
    copy_to_clipboard = False
    for opt, arg in opts:
        if opt in ("-d", "--dashless"):
            dashless = True
        elif opt in ("-o", "--override"):
            override = True
        elif opt in ("-m", "--maxchars"):
            max_characters = int(arg)
        elif opt in ("-c", "--copy"):
            copy_to_clipboard = True
        elif opt in ("-e", "--email"):
            email = arg
    return email, max_characters, dashless, override, copy_to_clipboard


def main(args):
    if len(args) < 2:
        raise Exception("Not enough arguments")
    request_type = args[0]
    if ["generate", "gen", "get", "make", "delete", "remove"].count(request_type) == 0:
        raise Exception("invalid request type")
    initialize_sheets_api()
    data = args[1]
    email, max_characters, dashless, override, copy_to_clipboard = get_options(args[2:])
    if request_type in ("generate", "gen"):
        password = to_password(data, max_characters, dashless)
        print_password(password, GET_MESSAGE, copy_to_clipboard)
    elif request_type == "get":
        row_data, _ = find_row_starting_with(data)
        password = get_password_from_row_data(row_data)
        print("Email/Username: %s" % row_data[EMAIL])
        print_password(password, GET_MESSAGE, copy_to_clipboard)
    elif request_type == "make":
        password = process_make_request(data, email, max_characters, dashless, override)
        print_password(password, MAKE_MESSAGE, copy_to_clipboard)
    elif request_type in  ("delete", "remove"):
        _, index = find_row_starting_with(data)
        delete_row(index)


if __name__ == '__main__':
    try:
        main(sys.argv[1:])
    except Exception as err:
        print(err)
