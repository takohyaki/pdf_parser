from django.shortcuts import render
from django.http import HttpResponse
from .forms import UploadFileForm
from .models import UploadedFile
from .tasks import extract_LOF_info, extract_app_info, add_row_to_excel
import os

def file_upload(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = UploadedFile(file=request.FILES['file'])
            uploaded_file.save()

            # After saving, process the file
            file_path = uploaded_file.file.path
            if 'Letter_of_Offer' in file_path:
                parsed_data = extract_LOF_info(file_path)
            else:  # Treat all other PDF files as application letters
                parsed_data = extract_app_info(file_path)

            # Check if parsing was successful
            if parsed_data:
                # Update the Excel file
                add_row_to_excel(parsed_data, file_path='./MRA_2023_data.xlsx')
                return render(request, 'fileuploader/upload_success.html')
            else:
                return HttpResponse("Failed to parse the file.", status=500)
    else:
        form = UploadFileForm()
    return render(request, 'fileuploader/upload.html', {'form': form})


