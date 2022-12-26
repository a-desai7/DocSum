from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']

def main():
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('drive', 'v3', credentials=creds)

        # Retrieve a list of documents
        query = "mimeType='application/vnd.google-apps.document'"
        results = service.files().list(q=query, pageSize=10, fields="nextPageToken, files(id, name)").execute()
        items = results.get("files", [])

        docs_service = build('docs', 'v1', credentials=creds)
        word_count = 0;

        # Iterate through the list of documents
        for item in items:
            # Use the `documents().get()` method to retrieve the content of the document as a text string
            document_id = item['id']
            document = docs_service.documents().get(documentId=document_id).execute()
            content = document.get('body').get('content')

            # Iterate through the list of text runs in the document
            for element in content:
                # Get the text of the text run
                text_run = element.get('textRun')
                if text_run:
                    # Get the text
                    text = text_run.get('content')

                    # Count the number of words in the text run
                    words = text.split()
                    num_words = len(words)

                    # Add the number of words in this text run to the running total
                    word_count += num_words

        print(f"Total number of words: {word_count}")
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()