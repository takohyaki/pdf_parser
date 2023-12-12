from django.conf import settings
from msal import ConfidentialClientApplication
import requests

client_id = settings.CLIENT_ID
client_secret = settings.CLIENT_SECRET
tenant_id = settings.TENANT_ID
drive_id = settings.DRIVE_ID
excel_file_path = settings.EXCEL_FILE_PATH

class MicrosoftGraphClient:
    def __init__(self, client_id, client_secret, tenant_id):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f'https://login.microsoftonline.com/{tenant_id}'
        self.scope = ['https://graph.microsoft.com/.default']
        self.access_token = None
        self.client = ConfidentialClientApplication(
            client_id,
            authority=self.authority,
            client_credential=client_secret,
        )

    def get_access_token(self):
        result = self.client.acquire_token_for_client(scopes=self.scope)
        if 'access_token' in result:
            self.access_token = result['access_token']
        return self.access_token

    def get_file(self, excel_file_path):
        if self.access_token is None:
            self.get_access_token()

        headers = {'Authorization': 'Bearer ' + self.access_token}
        file_url = f'https://graph.microsoft.com/v1.0/me/drive/root:/{excel_file_path}'
        response = requests.get(file_url, headers=headers)

        if response.status_code == 200:
            return response.json()
        else:
            return None
    
    def download_file(self, drive_id, excel_file_path):
        if self.access_token is None:
            self.get_access_token()

        headers = {'Authorization': 'Bearer ' + self.access_token}
        file_url = f'https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{excel_file_path}:/content'
        response = requests.get(file_url, headers=headers)

        if response.status_code == 200:
            return response.content
        else:
            return None
    
    def upload_file(self, drive_id, excel_file_path, file_content):
        if self.access_token is None:
            self.get_access_token()
        content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

        headers = {
            'Authorization': 'Bearer ' + self.access_token,
            'Content-Type': content_type  
        }
        upload_url = f'https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{excel_file_path}:/content'
        response = requests.put(upload_url, headers=headers, data=file_content)

        if response.status_code in [200, 201, 202]:
            return True
        else:
            print(f"Upload Failed: Status Code: {response.status_code}, Response: {response.text}")
            return False