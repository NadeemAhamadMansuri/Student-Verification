from flask import Flask, render_template, request, redirect
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from werkzeug.utils import secure_filename
import os
import uuid

# Firebase initialization
cred = credentials.Certificate("serviceAccountKey.json")
firebase_admin.initialize_app(cred)
db = firestore.client()

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# Load Excel student data
excel_file = "students.xlsx"
df = pd.read_excel(excel_file)

@app.route("/")
def index():
    # For now load first student row (can change for dynamic use)
    student_data = df.iloc[0].to_dict()
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

    # Merge data
    submitted_data = {**form_data, **uploaded_files}

    # Save to Firebase
    db.collection("student_verifications").add(submitted_data)

    # Save to Excel
    existing_df = pd.read_excel("submitted_data.xlsx") if os.path.exists("submitted_data.xlsx") else pd.DataFrame()
    new_df = pd.DataFrame([submitted_data])
    final_df = pd.concat([existing_df, new_df], ignore_index=True)
    final_df.to_excel("submitted_data.xlsx", index=False)

    return "Form submitted successfully!"

if __name__ == "__main__":
    app.run(debug=True)
