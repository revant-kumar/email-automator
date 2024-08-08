import pandas as pd
import os.path
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is created automatically
    # when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def send_email(service, user_id, to, subject, message_text, image_path):
    message = MIMEMultipart('related')
    message['to'] = to
    message['from'] = user_id
    message['subject'] = subject

    # Create the alternative part for plain text and HTML content
    msg_alternative = MIMEMultipart('alternative')
    message.attach(msg_alternative)

    # Create the plain-text and HTML version of your message
    text = message_text
    html = """
    <html>
      <body>
        <p>{}</p>
        <img src="cid:image1">
      </body>
    </html>
    """.format(message_text.replace('\n', '<br>'))

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')

    # Attach parts into message container.
    msg_alternative.attach(part1)
    msg_alternative.attach(part2)

    # Attach the image file
    with open(image_path, 'rb') as img_file:
        img = MIMEImage(img_file.read())
        img.add_header('Content-ID', '<image1>')
        message.attach(img)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    body = {'raw': raw}
    try:
        message = service.users().messages().send(userId=user_id, body=body).execute()
        print(f'Sent email to {to}')
    except HttpError as error:
        print(f'An error occurred: {error}')

def main():
    # Load the Excel sheet
    df = pd.read_excel('test.xlsx', header=None)

    creds = get_credentials()
    service = build('gmail', 'v1', credentials=creds)

    user_id = 'me'
    common_text = (
        #if any
    )

    image_path = 'outro.jpeg' # (optional)

    # Loop through the dataframe and send emails
    for index, row in df.iterrows():
        to_email = row[0]  
        receiver_name = row[1]  
        subject = row[2]  
        custom_paragraph = row[3] 

        email_body = (
            "Dear {},\n\n"
            "{}\n\n"
            "{}\n\nRegards,\n"
        ).format(receiver_name, custom_paragraph, common_text)

        send_email(service, user_id, to_email, subject, email_body, image_path)

if __name__ == '__main__':
    main()
