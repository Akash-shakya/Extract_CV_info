
#new
from fastapi import FastAPI, File, UploadFile, HTTPException
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from openpyxl import Workbook
import os
import re
import docx2txt
from PyPDF2 import PdfReader
import time

app = FastAPI()

def extract_contacts_from_pdf(pdf_path):
    contacts = {
        "texts": [],
        "emails": [],
        "phone_numbers": []
    }

    email_pattern = re.compile(r'[\w\.-]+@[\w\.-]+')
    phone_pattern = re.compile(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]')

    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        for page in reader.pages:
            text = page.extract_text()

            emails = email_pattern.findall(text)
            phone_numbers = phone_pattern.findall(text)

            contacts["texts"].append(text)
            contacts["emails"].extend(emails)
            contacts["phone_numbers"].extend(phone_numbers)

    return contacts

def extract_contacts_from_docx(docx_path):
    contacts = {
        "texts": [],
        "emails": [],
        "phone_numbers": []
    }

    email_pattern = re.compile(r'[\w\.-]+@[\w\.-]+')
    phone_pattern = re.compile(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]')

    text = docx2txt.process(docx_path)
    
    emails = email_pattern.findall(text)
    phone_numbers = phone_pattern.findall(text)

    contacts["texts"].append(text)
    contacts["emails"].extend(emails)
    contacts["phone_numbers"].extend(phone_numbers)
    
    return contacts

def extract_contacts_from_doc(doc_path):
    contacts = {
        "texts": [],
        "emails": [],
        "phone_numbers": []
    }

    email_pattern = re.compile(r'[\w\.-]+@[\w\.-]+')
    phone_pattern = re.compile(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]')

    with open(doc_path, 'rb') as file:
        # Try decoding the content using different encoding formats
        try:
            text = file.read().decode("utf-8")
        except UnicodeDecodeError:
            try:
                text = file.read().decode("latin-1")
            except UnicodeDecodeError:
                text = file.read().decode("utf-16")

    emails = email_pattern.findall(text)
    phone_numbers = phone_pattern.findall(text)

    contacts["texts"].append(text)
    contacts["emails"].extend(emails)
    contacts["phone_numbers"].extend(phone_numbers)
    
    return contacts


@app.post("/contacts")
async def extract_contacts_from_zip(file: UploadFile = File(...)):
    if not file.filename.endswith('.zip'):
        raise HTTPException(status_code=400, detail="Only ZIP files are allowed.")

    with TemporaryDirectory() as temp_dir:
        zip_path = os.path.join(temp_dir, file.filename)
        with open(zip_path, "wb") as buffer:
            buffer.write(await file.read())

        contacts = {
            "texts": [],
            "emails": [],
            "phone_numbers": []
        }
        wb = Workbook()
        ws = wb.active
        ws.append(["Filename", "Emails", "Phone Numbers", "Text"])

        with ZipFile(zip_path, 'r') as zip_ref:
            for filename in zip_ref.namelist():
                if filename.endswith(".pdf"):
                    pdf_path = os.path.join(temp_dir, filename)
                    zip_ref.extract(filename, path=temp_dir)
                    contacts = extract_contacts_from_pdf(pdf_path)
                elif filename.endswith(".docx"):
                    docx_path = os.path.join(temp_dir, filename)
                    zip_ref.extract(filename, path=temp_dir)
                    contacts = extract_contacts_from_docx(docx_path)
                elif filename.endswith(".doc"):
                    doc_path = os.path.join(temp_dir, filename)
                    zip_ref.extract(filename, path=temp_dir)
                    contacts = extract_contacts_from_doc(doc_path)
                else:
                    continue

                for email, phone, text in zip(contacts["emails"], contacts["phone_numbers"], contacts["texts"]):
                    ws.append([filename, email, phone, text])

        timestamp = time.strftime("%H%M%S")
        folder_name = "contacts"
        folder_path = os.path.join(os.getcwd(), folder_name)
        if not os.path.exists(folder_path):
            # Create the folder only if it doesn't exist
            os.mkdir(folder_path)

        excel_filename = f"contactsCV_{timestamp}.xlsx"
        excel_file_path = os.path.join(folder_path, excel_filename)
        wb.save(excel_file_path)

        return {"excel_file_path": excel_file_path}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)