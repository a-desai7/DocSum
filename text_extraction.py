# Copyright 2023 Aayush Desai
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# Adapted from Google Workspace: https://developers.google.com/docs/api/samples/extract-text#python
# Accessed December 25, 2022

"""
Recursively extracts the text from each Google Doc in a user's Google Drive folder.
Prints the total number of documents and words.
"""
import time
import numpy
import googleapiclient.discovery as discovery
from httplib2 import Http
from oauth2client import client
from oauth2client import file
from oauth2client import tools


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
DISCOVERY_DOC = 'https://docs.googleapis.com/$discovery/rest?version=v1'
DOCUMENT_ID = '1IEK_KWElKX3PohdtFlU_s7KiD6sje_cerVw3Esw814k'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth 2.0 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    store = file.Storage('token.json')
    credentials = store.get()

    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        credentials = tools.run_flow(flow, store)
    return credentials

def read_paragraph_element(element):
    """Returns the text in the given ParagraphElement.

        Args:
            element: a ParagraphElement from a Google Doc.
    """
    text_run = element.get('textRun')
    if not text_run:
        return ''
    return text_run.get('content')


def read_structural_elements(elements):
    """Recurses through a list of Structural Elements to read a document's text where text may be
        in nested elements.

        Args:
            elements: a list of Structural Elements.
    """
    text = ''
    try:
        for value in elements:
            if 'paragraph' in value:
                elements = value.get('paragraph').get('elements')
                for elem in elements:
                    text += read_paragraph_element(elem)
            elif 'table' in value:
                # The text in table cells are in nested Structural Elements and tables may be
                # nested.
                table = value.get('table')
                for row in table.get('tableRows'):
                    cells = row.get('tableCells')
                    for cell in cells:
                        text += read_structural_elements(cell.get('content'))
            elif 'tableOfContents' in value:
                # The text in the TOC is also in a Structural Element.
                toc = value.get('tableOfContents')
                text += read_structural_elements(toc.get('content'))
        return text
    except:
        return text


def main():
    """Uses the Docs API to authenticate user information."""
    credentials = get_credentials()
    http = credentials.authorize(Http())
    docs_service = discovery.build(
        'docs', 'v1', http=http, discoveryServiceUrl=DISCOVERY_DOC)

    drive_service = build('drive', 'v3', credentials=credentials)

    """ Retrieves a list of documents in the user's Drive folder"""
    query = "mimeType='application/vnd.google-apps.document'"
    results = drive_service.files().list(q=query, pageSize=1000, fields="nextPageToken, files(id, name)").execute()
    items = results.get("files", [])

    start = time.time()
    """Iterates through each document and adds to the running sum"""
    word_total = 0
    doc_total = 0
    lengthiest_doc = (0, "")
    least_lengthy_doc = (float('inf'), "")
    num_words_list = []
    for item in items:
        document_id = item['id']
        doc = docs_service.documents().get(documentId=document_id).execute()
        doc_content = doc.get('body').get('content')
        doc_title = doc.get('title')

        words = read_structural_elements(doc_content)
        num_words = len(words.split())
        num_words_list.append(num_words)
        if (num_words > lengthiest_doc[0]):
            lengthiest_doc = (num_words, doc_title)
        if (num_words < least_lengthy_doc[0]):
            least_lengthy_doc = (num_words, doc_title)
        word_total += num_words
        doc_total += 1

    minimum = numpy.min(num_words_list)
    maximum = numpy.max(num_words_list)
    mean = round(numpy.mean(num_words_list), 1)
    median = round(numpy.median(num_words_list), 1)
    std = round(numpy.std(num_words_list), 1)

    print(f"\nGlobal word total: {word_total}")
    print(f"Global document total: {doc_total}")
    
    print(f"Minimum: {least_lengthy_doc}")
    print(f"Maximum: {lengthiest_doc}")
    print(f"Mean: {mean}")
    print(f"Median: {median}")
    print(f"Standard deviation: {std}")
    end = time.time()
    elapsed_time = (int) (end - start)
    print('Execution time:', elapsed_time, 'seconds')

if __name__ == '__main__':
    main()
