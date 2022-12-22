from __future__ import print_function

import os.path
import io

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/documents']  


def main():
    """Shows basic usage of the Drive v3 API.
    Prints the total number of words a user has written in their Google Docs.
    """
    creds = None
    word_count = 0;
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

        # Call the Drive v3 API
        results = service.files().list(
            pageSize=10, fields="nextPageToken, files(id, name)").execute()
        items = results.get('files', [])

        if not items:
            print('No files found.')
            return
        print('Files:')

        # Iterate through the list of items, downloading the content of each as a stream
        for item in items:
            file_id = item['id']
            request = service.files().export(fileId=file_id, mimeType='text/plain')
            file = io.BytesIO()
            downloader = MediaIoBaseDownload(file, request)

            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print(f'Download {int(status.progress() * 100)}.')
            file.seek(0)
            content = file.read().decode()

            # Use Python's `len()` function to count the number of words in the text string
            words = content.split()
            num_words = len(words)

            # Add the number of words in this file to the running total
            word_count += num_words
    except HttpError as error:
        # TODO(developer) - Handle errors from drive API.
        print(f'An error occurred: {error}')
    
    print(f"Total number of words: {word_count}")

if __name__ == '__main__':
    main()