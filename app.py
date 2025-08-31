from flask import Flask, render_template, request, redirect
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
firebase_key = os.environ.get("service_account_json")  # Render environment variable
if not firebase_key:
    raise Exception("service_account_json environment variable not set")

# escaped newline को actual newline में बदलें ताकि PEM सही आए
firebase_key = firebase_key.replace('\\n', '\n')

cred_dict = json.loads(firebase_key)
cred = credentials.Certificate(cred_dict)
firebase_admin.initialize_app(cred)
db = firestore.client()


# -------------------------------
# Flask setup
# -------------------------------
app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)


# -------------------------------
# Load Excel student data
# -------------------------------
excel_file = os.environ.get("STUDENT_EXCEL", "students.xlsx")  # Render env variable for Excel name
df = pd.read_excel(excel_file)


# -------------------------------
# Routes
# -------------------------------
@app.route("/")
def index():
    student_data = df.iloc[0].to_dict()  # Load first student row (dynamic can be added later)
    return render_template("index.html", student_data=student_data)


@app.route("/submit", methods=["POST"])
def submit():
    form_data = request.form.to_dict()

    # Handle file uploads
    uploaded_files = {}
    for file_key in request.files:
        file = request.files[file_key]
        if file and file.filename != "":
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
    excel_path = os.environ.get("SUBMITTED_EXCEL", "submitted_data.xlsx")  # Render env variable
    existing_df = pd.read_excel(excel_path) if os.path.exists(excel_path) else pd.DataFrame()
    new_df = pd.DataFrame([submitted_data])
    final_df = pd.concat([existing_df, new_df], ignore_index=True)
    final_df.to_excel(excel_path, index=False)

    return "Form submitted successfully!"


# -------------------------------
# Run app
# -------------------------------
if __name__ == "__main__":
    app.run(debug=True)
