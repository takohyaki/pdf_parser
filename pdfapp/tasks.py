import os
import re
import fitz
import pandas as pd
import openpyxl
import io
from msal import ConfidentialClientApplication
from django.conf import settings

reference_id_pattern = re.compile(r"Ref ID:\s([0-9A-Z]+)\s*\n\s*([0-9]{2} [A-Za-z]{3} 2023)?")
application_date_pattern = re.compile(r"application dated ([0-9]{2} [A-Za-z]{3} 2023)")

def extract_LOF_info(pdf_path):

    # file name with extension
    file_name = os.path.basename(pdf_path)

    if 'VTF' in file_name:
        package = 'MRA1'
    elif 'BM' in file_name:
        package = 'MRA2'
    elif 'MRA3' in pdf_path:
        package = 'MRA3'
    else:
        package = 'Unknown'

    # Open the PDF file
    doc = fitz.open(pdf_path)

    # Extract text from all pages
    text = "\n".join([page.get_text() for page in doc])
    # print(text)

    reference_id = re.search(reference_id_pattern, text)
    application_date = re.search(application_date_pattern, text)
    
    reference_id_match = re.search(reference_id_pattern, text)
    if reference_id_match:
        reference_id = reference_id_match.group(1)
        # Check if the date was captured
        application_approved_date = reference_id_match.group(2) if reference_id_match.group(2) else "Not Found"
    else:
        reference_id = "Not Found"
        application_approved_date = "Not Found"

    # Find 'Annex 1' section
    annex1_text = re.search(r"Annex 1 – Details of Eligible Expenses for the Project.*?(?=Annex|$)", text, re.DOTALL)
    if annex1_text:
        annex1_text = annex1_text.group(0)

        # Parsing within Annex 1 to find the values after '[a]'
        match = re.search(r"\[a\][\s\S]*?(\d+)\s+([A-Z]+)\s+([\d,\.]+)", annex1_text)
        if match:
            level_of_support = match.group(1)
            billing_currency = match.group(2)
            value = match.group(3)
            # Finding the last occurrence of the value in the sequence for 'Qualifying Cost'
            qualifying_cost = match.group(3).split()[-1]
        else:
            level_of_support = billing_currency = value = qualifying_cost = "Not Found"
    else:
        level_of_support = billing_currency = value = qualifying_cost = "Not Found"

    # Extracted information
    extracted_info = {
        "Reference ID": reference_id,
        "Application Approved Date": application_approved_date,
        "Value": value,
        "Application Date": application_date.group(1) if application_date else "Not Found",
        "Package": package,
        "Level of Support (%)": level_of_support,
        "Billing Currency": billing_currency,
        "Qualifying Cost": qualifying_cost
    }

    return extracted_info

# Precompile regular expressions
project_title_pattern = re.compile(r"Project Title\s*\n(.*?)\n", re.DOTALL)
project_start_date_pattern = re.compile(r"Project Title.*?Start Date\s*\n(.*?)\n", re.DOTALL)
project_end_date_pattern = re.compile(r"Project Title.*?End Date\s*\n(.*?)\n", re.DOTALL)
application_type_pattern = re.compile(r"Target Market\s*\n(.*?)(?=\n)", re.DOTALL)
ref_id_pattern = re.compile(r"Ref ID:\s*(.*)")
company_name_pattern = re.compile(r"Registered Company Name\s*\n(.*)")

def extract_app_info(file_path):
    try:
        with fitz.open(file_path) as doc:
            text = "\n".join([page.get_text() for page in doc])

        project_title_match = project_title_pattern.search(text)
        project_start_date_match = project_start_date_pattern.search(text)
        project_end_date_match = project_end_date_pattern.search(text)
        application_type_match = application_type_pattern.search(text)
        ref_id_match = ref_id_pattern.search(text)
        company_name_match = company_name_pattern.search(text)

        extracted_info = {
            "Company": company_name_match.group(1).strip() if company_name_match else "Not Found",
            "Project Title": project_title_match.group(1).strip() if project_title_match else "Not Found",
            "Project Start Date": project_start_date_match.group(1).strip() if project_start_date_match else "Not Found",
            "Project End Date": project_end_date_match.group(1).strip() if project_end_date_match else "Not Found",
            "Application Type": application_type_match.group(1).strip() if application_type_match else "Not Found",
            "Reference ID": ref_id_match.group(1).strip() if ref_id_match else "Not Found"
        }

        return extracted_info
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return None

def add_row_to_excel(graph_client, data):
    file_content = graph_client.download_file(settings.DRIVE_ID, settings.EXCEL_FILE_PATH)
    if file_content is None:
        return "Error: File not found or unable to download."

    try:
        df = pd.read_excel(io.BytesIO(file_content))
    except FileNotFoundError:
        df = pd.DataFrame(columns=data.keys())

    ref_id = data.get('Reference ID', 'Default Value')

    if ref_id in df['Reference ID'].values:
        index = df.index[df['Reference ID'] == ref_id].tolist()[0]
        for key in data.keys():
            df.at[index, key] = data[key]
    else:
        new_row = pd.DataFrame([data])
        df = pd.concat([df, new_row], ignore_index=True)

    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    upload_success = graph_client.upload_file(settings.DRIVE_ID, settings.EXCEL_FILE_PATH, output.read())
    return "File updated successfully." if upload_success else "Error: Unable to upload the file."

