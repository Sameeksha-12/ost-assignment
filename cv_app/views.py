from django.shortcuts import render
import pythoncom
import win32com.client
import docx2txt
import re
import os
import tempfile
import fitz
from .forms import FileUploadForm
import json
from django.http import HttpResponse
from openpyxl import Workbook



def upload_files(request):
    if request.method == 'POST':
        form = FileUploadForm(request.POST, request.FILES)
        if form.is_valid():
            extracted_data_list = []
            for file in request.FILES.getlist('files'):
                # Process each uploaded file
                extracted_data = process_file(file)
                extracted_data_list.append(extracted_data)
            return render(request, 'display_data.html', {'extracted_data_list': extracted_data_list})
    else:
        form = FileUploadForm()
    return render(request, 'upload_files.html', {'form': form})

def convert_doc_to_docx(doc_file):
    try:
        # Initialize COM
        pythoncom.CoInitialize()

        # Create a temporary directory to store the file
        temp_dir = tempfile.mkdtemp()

        # Save the uploaded DOC file to the temporary directory
        doc_file_path = os.path.join(temp_dir, doc_file.name)
        with open(doc_file_path, 'wb+') as destination:
            for chunk in doc_file.chunks():
                destination.write(chunk)

        # Create a new instance of Word application
        word = win32com.client.Dispatch("Word.Application")
        
        # Create a temporary DOCX file path
        temp_docx_file = os.path.splitext(doc_file.name)[0] + ".docx"
        temp_docx_file_path = os.path.join(temp_dir, temp_docx_file)
        
        # Save the document in DOCX format
        word.Documents.Open(doc_file_path)
        word.ActiveDocument.SaveAs(temp_docx_file_path, FileFormat=16)
        
        # Close Word application
        word.Quit()

        return temp_docx_file_path
    finally:
        # Uninitialize COM
        pythoncom.CoUninitialize()

def process_file(file):
    # Initialize variables to store extracted data
    email = None
    contact_number = None
    overall_text = ""

    # Check if the uploaded file is a DOCX file
    if file.name.endswith('.docx'):
        overall_text = docx2txt.process(file)

    # Check if the uploaded file is a DOC file
    elif file.name.endswith('.doc'):
        # Convert DOC file to DOCX format
        docx_file_path = convert_doc_to_docx(file)
        # Process the converted DOCX file
        overall_text = docx2txt.process(docx_file_path)
        # Delete the temporary DOCX file
        os.remove(docx_file_path)

    # Check if the uploaded file is a PDF file
    elif file.name.endswith('.pdf'):
        # Open the PDF file
        pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            overall_text += page.get_text()

    # Print extracted text for debugging
    # print("Extracted Text:", overall_text)

    # Extract email and contact number from text
    email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    contact_regex = r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})'
    email_match = re.search(email_regex, overall_text)
    contact_match = re.search(contact_regex, overall_text)

    # Assign extracted email and contact number if found
    if email_match:
        email = email_match.group(0)
    if contact_match:
        contact_number = contact_match.group(0)

    # Return the extracted data as a dictionary
    extracted_data = {
        'email': email,
        'contact_number': contact_number,
        'overall_text': overall_text
    }
    return extracted_data

import json
from django.http import HttpResponse
from openpyxl import Workbook

def download_excel(request):
    if request.method == 'POST':
        try:
            # Retrieve extracted_data_list from POST data
            extracted_data_list_json = request.POST.get('extracted_data_list')
            extracted_data_list = json.loads(extracted_data_list_json)

            # Create a new Excel workbook
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Extracted Data"

            # Add headers
            worksheet['A1'] = 'Email'
            worksheet['B1'] = 'Contact Number'
            worksheet['C1'] = 'Overall Text'

            # Add data from extracted_data_list
            for idx, extracted_data in enumerate(extracted_data_list, start=2):
                worksheet[f'A{idx}'] = extracted_data.get('email', '')
                worksheet[f'B{idx}'] = extracted_data.get('contact_number', '')
                worksheet[f'C{idx}'] = extracted_data.get('overall_text', '')

            # Create a response object with the Excel file
            response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="extracted_data.xlsx"'

            # Save the Excel workbook to the response
            workbook.save(response)

            return response
        except Exception as e:
            return HttpResponse(f"An error occurred: {str(e)}", status=500)

    return HttpResponse("Invalid request", status=400)
