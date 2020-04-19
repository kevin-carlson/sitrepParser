from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient.discovery import build
from apiclient import errors
from httplib2 import Http
from oauth2client import file, client, tools
from email.mime.text import MIMEText
from base64 import urlsafe_b64encode
import base64

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/gmail.send']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1IpEP3m4p0LQlD82DSo9UTSjnrTK0S8qv1A5MURd7hhU'
SHEET_ID = '439508074'
SAMPLE_RANGE_NAME = "'Form Responses 1'!A2:J"

def main():
    #CREATE CREDENTIALS
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    #service for google sheets:
    service = build('sheets', 'v4', credentials=creds)
    #service for gmail
    service_gmail = build('gmail', 'v1', credentials=creds)


    #Call sheets api
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        #initialize empty string for the messages
        mit1_text = ''
        mit11_text = ''
        mit12_text = ''
        mit13_text = ''

        mit2_text = ''
        mit21_text = ''
        mit22_text = ''
        mit23_text = ''

        bu1_text = ''
        bu11_text = ''
        bu12_text = ''
        bu13_text = ''

        bu2_text = ''
        bu21_text = ''
        bu22_text = ''
        bu23_text = ''

        mit_costaff = ''
        bu_costaff = ''
        s3 = ''
        battstaff = ''
        marines = ''
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[4]))
            #Must do elifs for other platoons/groups
            #MIT Platoon 1
            if row[3]=="MIT 1":
                #MIT 1st Platoon
                mit1_text = mit1_text + "<p><b>"+row[2]+"</b></p>"
                mit1_message = create_message("kccscout1@gmail.com","kcarlson@mit.edu", "Test email data from sheet",mit1_text)
                print("send to ramirez")
                if row[4]=="1":
                    #parse data to send to Mendez
                    mit11_text = mit11_text + "<p><b>"+row[2]+"</b></p>"
                if row[4]=="2":
                    #parse data to send to Edelman
                    mit12_text = mit12_text + "<p><b>"+row[2]+"</b></p>"
                if row[4]=="3":
                    #parse data to send to Worthley
                    mit13_text = mit13_text + "<p><b>"+row[2]+"</b></p>"
        #mit 1
        mit1_message = create_message("kccscout1@gmail.com","kcarlson@mit.edu", "MIT 1 SITREPS",mit1_text)
        send_ramirez = send_message(service_gmail,"me",mit1_message)
        print("sent mit 1")
        #mit 1-1
        mit11_message = create_message("kccscout1@gmail.com","kcarlson@mit.edu", "MIT 1-1 SITREPS",mit11_text)
        send_mendez = send_message(service_gmail,"me",mit11_message)
        print("sent mit 1-1")
        #mit 1-2
        mit12_message = create_message("kccscout1@gmail.com","kcarlson@mit.edu", "MIT 1-2 SITREPS",mit12_text)
        send_edelman = send_message(service_gmail,"me",mit12_message)
        print("sent mit 1-2")
        #mit 1-2
        mit13_message = create_message("kccscout1@gmail.com","kcarlson@mit.edu", "MIT 1-3 SITREPS",mit13_text)
        send_edelman = send_message(service_gmail,"me",mit13_message)
        print("sent mit 1-3")

def create_message(sender, to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEText(message_text, 'html')
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def send_message(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    return message
  except errors.HttpError, error:
    print ('An error occurred:')


if __name__ == '__main__':
    main()
