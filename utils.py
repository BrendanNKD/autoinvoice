import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

class Utils:
    SCOPES = ["https://www.googleapis.com/auth/drive"]

    @staticmethod
    def load_credentials():
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", Utils.SCOPES)
            if creds and creds.valid:
                return creds

        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", Utils.SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
        return creds

    @staticmethod
    def get_folder(foldername, service):
        folder_query = f"name = '{foldername}' and mimeType = 'application/vnd.google-apps.folder'"
        folder_results = service.files().list(q=folder_query, fields="files(id, name)").execute()
        folder_items = folder_results.get("files", [])
        return folder_items
