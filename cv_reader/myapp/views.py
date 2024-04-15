import os
import zipfile
import re
import PyPDF2
from openpyxl import Workbook
from django.shortcuts import render
from django.conf import settings
from django.http import HttpResponseRedirect, HttpResponse
from django.core.files.storage import FileSystemStorage

def extract_data_from_pdf(pdf_file):
    text = ""
    reader = PyPDF2.PdfReader(pdf_file)
    for page_num in range(len(reader.pages)):
        text += reader.pages[page_num].extract_text()
    return text

def find_contact_info(text):
    phone_regex = r'(\+\d{1,3}\s?)?(\(?\d{3}\)?[\s.-]?)?(\d{10})'
    email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

    phone_numbers = re.findall(phone_regex, text)
    emails = re.findall(email_regex, text)

    mobile = next((num[2] for num in phone_numbers if len(num[2]) == 10), None)
    email = emails[0] if emails else None

    return mobile, email

def process_resumes(zip_file_path, output_excel_path):
    wb = Workbook()
    ws = wb.active
    ws.append(['Mobile Number', 'Email'])

    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        for filename in zip_ref.namelist():
            if filename.endswith('.pdf'):
                with zip_ref.open(filename) as f:
                    text = extract_data_from_pdf(f)
                    mobile, email = find_contact_info(text)
                    ws.append([mobile, email])

    wb.save(output_excel_path)

def home(request):
    if request.method == 'POST' and request.FILES.get('cv_zip'):
        uploaded_file = request.FILES['cv_zip']
        fs = FileSystemStorage()
        zip_file_name = uploaded_file
        zip_file_path = fs.save(zip_file_name, uploaded_file)

        extract_dir = os.path.join(settings.MEDIA_ROOT, 'extracted_cvs')
        os.makedirs(extract_dir, exist_ok=True)
        output_excel_path = os.path.join(extract_dir, 'resumes_data.xlsx')

        process_resumes(zip_file_name, output_excel_path)

        with open(output_excel_path, 'rb') as excel_file:
            response = HttpResponse(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename=resumes_data.xlsx'
            try:
                os.remove(zip_file_path)
                print("Uploaded zip file deleted successfully.")
            except Exception as e:
                print("Error deleting uploaded zip file:", e)
            return response

    return render(request, 'home.html')
