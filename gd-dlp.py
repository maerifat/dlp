from __future__ import print_function
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def search_file():
    """Search file in drive location"""
    # Path to the service account key file downloaded manually
    key_path = "rahul.json"

    # Load credentials from the key file
    creds = None
    if os.path.exists(key_path):
        creds = service_account.Credentials.from_service_account_file(
            key_path, scopes=["https://www.googleapis.com/auth/drive"],
            subject="xxyyzzrupifi.com"
        )

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)
        folders = []
        page_token = None
        while True:
            # pylint: disable=maybe-no-member
            response = service.files().list(q="visibility='anyoneCanFind' and trashed = false",
                                            spaces='drive',
                                            fields='nextPageToken, files(id, name, mimeType, owners, webViewLink)',
                                            pageToken=page_token).execute()

            for file in response.get('files', []):
                # Process file/folder
                owners = file.get('owners', [])
                owner_emails = [owner.get('emailAddress') for owner in owners]
                if file.get('mimeType') == 'application/vnd.google-apps.folder':
                    print(F'Found folder: {file.get("name")}, {file.get("id")}, shared with: {", ".join(owner_emails)}')
                    folders.append(file)
                else:
                    print(F'Found file: {file.get("name")}, {file.get("id")}, shared with: {", ".join(owner_emails)}, view link: {file.get("webViewLink")}')

            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


search_file()

#export GOOGLE_APPLICATION_CREDENTIALS=`pwd`/key3.json
