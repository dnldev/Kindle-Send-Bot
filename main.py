import json
import pickle
import os
import os.path
import base64
import mimetypes
import sys
import googleapiclient.errors as errors
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from os import listdir


def epub2mobi(fromdir, todir, ignore_if=None):
    """Look for .epub files in fromdir, convert them to .mobi and store 
    them in the flat directory todir unless their path includes any string
    present in the list ignore_if. 
    Requires ebook-convert, coming from calibre (http://calibre-ebook.com). 
    Go to Preferences, select Miscellaneous in Advanced, and click the 
    "Install command line tools" button.
    """
    if not os.path.exists(todir):
        os.makedirs(todir)
    for root, dirs, files in os.walk(fromdir):
        ignore = False
        if ignore_if is not None:
            ignore = reduce(lambda a, b: a or b,
                            [ig in root for ig in ignore_if])
        if not ignore:
            for fl in files:
                nm, ext = os.path.splitext(fl)
                if ext == '.epub':
                    mobi = os.path.join(todir, nm + '.mobi')
                    if not os.path.exists(mobi):
                        os.system('ebook-convert ' +
                                  os.path.join(root, fl) + ' ' + mobi)


def create_message_with_attachment(
        sender, to, subject, message_text, files):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      message_text: The text of the email message.
      files: List of paths to the files to be attached.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    for file in files:
        content_type, encoding = mimetypes.guess_type(file)

        if content_type is None or encoding is not None:
            content_type = 'application/octet-stream'

        main_type, sub_type = content_type.split('/', 1)
        with open(file, 'rb') as fp:
            msg = MIMEBase(main_type, sub_type)
            msg.set_payload(fp.read())

        filename = os.path.basename(file)
        msg.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


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
        print('Message Id: %s' % message['id'])
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)
    # except BrokenPipeError as error:
    #     print('An error occurred: %s' % error)

    #     devnull = os.open(os.devnull, os.O_WRONLY)
    #     os.dup2(devnull, sys.stdout.fileno())
    #     sys.exit(1)


def create_service():
    SCOPES = ['https://www.googleapis.com/auth/gmail.send']

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

    service = build('gmail', 'v1', credentials=creds)
    return service


def main():
    service = create_service()

    sender = ""
    kindle = ""
    with open('information.json') as information:
        data = json.load(information)
        sender = data["email"]
        kindle = data["kindle-address"]

    epub2mobi("epubs", "mobis")

    files = list(map(lambda file: "mobis/" + file, listdir("mobis")))

    message = create_message_with_attachment(
        sender, kindle, "Mobi Files", "", files)

    send_message(service, sender, message)


if __name__ == '__main__':
    main()
