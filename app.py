from flask import Flask, render_template, request
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from werkzeug.utils import secure_filename
import os
import uuid
import json

# -------------------------------
# Firebase initialization
# -------------------------------
cred = None

if os.path.exists("service_account.json"):
    # Local JSON file
    cred = credentials.Certificate("service_account.json")
else:
    # Environment variable (Render)
    firebase_key = os.environ.get("service_account_json")
    if not firebase_key:
        raise Exception("service_account_json environment variable not set")
    try:
        firebase_key = firebase_key.replace('\\n', '\n')
        cred_dict = json.loads(firebase_key)
        cred = credentials.Certificate(cred_dict)
    except Exception as e:
        raise Exception(f"Failed to load Firebase credentials: {e}")

firebase_admin.initialize_app(cred)
db = firestore.client()

# -------------------------------
# Flask setup
# -------------------------------
app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# -------------------------------
# Load Excel student data
# -------------------------------
excel_file = os.environ.get("STUDENT_EXCEL", "students.xlsx")
try:
    df = pd.read_excel(excel_file)
except FileNotFoundError:
    df = pd.DataFrame()  # Empty if no file found

# -------------------------------
# Routes
# -------------------------------
@app.route("/")
def index():
    if df.empty:
        student_data = {}
    else:
        student_data = df.iloc[0].to_dict()
    return render_template("index.html", student_data=student_data)

@app.route("/submit", methods=["POST"])
def submit():
    form_data = request.form.to_dict()

    # Handle file uploads
    uploaded_files = {}
    for file_key in request.files:
        file = request.files[file_key]
        if file and file.filename:
            filename = secure_filename(f"{uuid.uuid4()}_{file.filename}")
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)
            uploaded_files[file_key] = filepath
        else:
            uploaded_files[file_key] = None

    # Merge form data and uploaded files
    submitted_data = {**form_data, **uploaded_files}

    # Save to Firebase
    db.collection("student_verifications").add(submitted_data)

    # Save to Excel
    excel_path = os.environ.get("SUBMITTED_EXCEL", "submitted_data.xlsx")
    try:
        existing_df = pd.read_excel(excel_path)
    except FileNotFoundError:
        existing_df = pd.DataFrame()
    new_df = pd.DataFrame([submitted_data])
    final_df = pd.concat([existing_df, new_df], ignore_index=True)
    final_df.to_excel(excel_path, index=False)

    return "Form submitted successfully!"

# -------------------------------
# Run app
# -------------------------------
if __name__ == "__main__":
    app.run(debug=True)
