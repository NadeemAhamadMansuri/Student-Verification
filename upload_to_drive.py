import os
import json
import firebase_admin
from firebase_admin import credentials, firestore
from flask import Flask, render_template, request, redirect
import pandas as pd
import openpyxl
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.service_account import Credentials

app = Flask(__name__)

# ---------------- Google Drive Auth ----------------
SERVICE_ACCOUNT_INFO = json.loads(os.environ.get("GOOGLE_SERVICE_ACCOUNT"))
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")
EXCEL_FILE_NAME = os.environ.get("EXCEL_FILE_NAME", "submitted_data.xlsx")

creds = Credentials.from_service_account_info(
    SERVICE_ACCOUNT_INFO,
    scopes=["https://www.googleapis.com/auth/drive.file"]
)
drive_service = build('drive', 'v3', credentials=creds)

# ---------------- Firebase Auth ----------------
if not firebase_admin._apps:
    cred = credentials.Certificate(SERVICE_ACCOUNT_INFO)
    firebase_admin.initialize_app(cred)
db = firestore.client()

# ---------------- Routes ----------------
@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    student_name = request.form['student_name']
    student_class = request.form['student_class']
    father_income = request.form['father_income']
    mother_income = request.form['mother_income']
    total_income = int(father_income) + int(mother_income)

    # Handle uploaded file
    uploaded_file = request.files['file']
    file_link = None
    if uploaded_file:
        local_path = os.path.join("uploads", uploaded_file.filename)
        uploaded_file.save(local_path)

        file_metadata = {
            "name": uploaded_file.filename,
            "parents": [DRIVE_FOLDER_ID]
        }
        media = MediaFileUpload(local_path, resumable=True)
        uploaded = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink"
        ).execute()
        file_link = uploaded.get("webViewLink")

    # Save to Firestore
    data = {
        "student_name": student_name,
        "student_class": student_class,
        "father_income": father_income,
        "mother_income": mother_income,
        "total_income": total_income,
        "file_link": file_link
    }
    db.collection("students").add(data)

    # Save to Excel
    if not os.path.exists(EXCEL_FILE_NAME):
        df = pd.DataFrame(columns=["Student Name", "Class", "Father Income", "Mother Income", "Total Income", "File Link"])
        df.to_excel(EXCEL_FILE_NAME, index=False)

    df = pd.read_excel(EXCEL_FILE_NAME)
    new_row = pd.DataFrame([{
        "Student Name": student_name,
        "Class": student_class,
        "Father Income": father_income,
        "Mother Income": mother_income,
        "Total Income": total_income,
        "File Link": file_link
    }])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(EXCEL_FILE_NAME, index=False)

    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)
