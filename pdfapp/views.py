from django.shortcuts import render, redirect
from django.http import HttpResponse
from .forms import UploadFileForm
# from .models import UploadedFile
from .tasks import extract_LOF_info, extract_app_info, add_row_to_excel
import os
from .microsoft_graph_client import MicrosoftGraphClient
from django.conf import settings
import requests

def index(request):
    return render(request, 'index.html')

def file_upload(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            file_content = file.read()  # Read the file content

            if 'Letter_of_Offer' in file.name:
                parsed_data = extract_LOF_info(file_content, file.name)
            else:
                parsed_data = extract_app_info(file_content, file.name)

            if parsed_data:
                graph_client = MicrosoftGraphClient(settings.CLIENT_ID, settings.CLIENT_SECRET, settings.TENANT_ID)
                result = add_row_to_excel(graph_client, parsed_data)

                if result == "File updated successfully.":
                    return render(request, 'fileuploader/upload_success.html', {
                        'reference_id': parsed_data.get('Reference ID', 'Unknown'),
                        'upload_success': True
                    })
                else:
                    return HttpResponse(result, status=500)
            else:
                return HttpResponse("Failed to parse the file.", status=500)
    else:
        form = UploadFileForm()
    return render(request, 'fileuploader/upload.html', {'form': form})

client_id = settings.CLIENT_ID
client_secret = settings.CLIENT_SECRET
tenant_id = settings.TENANT_ID
drive_id = settings.DRIVE_ID
excel_file_path = settings.EXCEL_FILE_PATH

def get_microsoft_file(self, drive_id, file_path):
    if self.access_token is None:
        self.get_access_token()

    headers = {'Authorization': 'Bearer ' + self.access_token}
    file_url = f'https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{excel_file_path}'
    response = requests.get(file_url, headers=headers)

    if response.status_code == 200:
        return response.json()
    else:
        return None

def home(request):
    return render(request, 'index.html')