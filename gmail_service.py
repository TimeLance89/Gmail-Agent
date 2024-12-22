import os
import base64
from email.parser import BytesParser
from email.policy import default

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def get_gmail_service():
    """
    Erstellt einen Gmail-Service mithilfe der OAuth2-Credentials.
    Gibt ein `googleapiclient.discovery.Resource`-Objekt zurück.
    """
    creds = None
    token_file = 'token.json'
    client_secret_file = 'client_secret.json'  # aus der Google Cloud Console
    
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
    
    service = build('gmail', 'v1', credentials=creds)
    return service


def list_emails(service, max_results=10):
    """
    Ruft die letzten `max_results` E-Mails aus dem Posteingang ab und gibt sie als Liste zurück.
    """
    results = service.users().messages().list(
        userId='me',
        labelIds=['INBOX'],
        maxResults=max_results
    ).execute()
    
    messages = results.get('messages', [])
    email_list = []
    
    for msg in messages:
        msg_data = service.users().messages().get(
            userId='me',
            id=msg['id'],
            format='raw'
        ).execute()
        
        # raw decodieren und parsen
        raw_data = base64.urlsafe_b64decode(msg_data['raw'].encode('ASCII'))
        email_message = BytesParser(policy=default).parsebytes(raw_data)
        
        subject = email_message.get('Subject', '')
        from_addr = email_message.get('From', '')
        snippet = msg_data.get('snippet', '')
        
        email_list.append({
            'subject': subject,
            'from': from_addr,
            'snippet': snippet
        })
    
    return email_list
