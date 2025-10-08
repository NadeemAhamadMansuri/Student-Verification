from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# -----------------------
# EXCEL FILES
# -----------------------
STUDENTS_FILE = "students.xlsx"
SUBMITTED_FILE = "submitted_data.csv"

# Create submitted_data.xlsx if it doesn't exist
if not os.path.exists(SUBMITTED_FILE):
    df = pd.DataFrame()
    df.to_excel(SUBMITTED_FILE, index=False)

# -----------------------
# GMAIL SETTINGS
# -----------------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "earthinfran.a@gmail.com"       # apna Gmail daalo
SENDER_PASSWORD = "aqybluvsaptonrvi"    # yaha apna Gmail App Password daalo
RECEIVER_EMAIL = "nadeemahamadmansuri@gmail.com"     # jis Gmail pe files aayengi

def send_email_with_files(subject, body, filepaths):
    try:
        logger.info("Preparing email...")  # Log email start
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = RECEIVER_EMAIL
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        # Attach all files
        for filepath in filepaths:
            with open(filepath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(filepath)}"
                )
                msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
        server.quit()
        logger.info("Email sent successfully")  # Log email success
        return True 
    except Exception as e:
        logger.error(f"Email send error: {str(e)}", exc_info=True)  # Email error log
        return f"Email send error: {str(e)}"


# -----------------------
# ROUTES
# -----------------------
@app.route('/', methods=['GET', 'POST'])
def index():
    logger.info("Index route called")  # Logging start
    error = None
    if request.method == 'POST':
        try:
            admission_no = str(request.form.get('admission_no')).strip()
            logger.info(f"Searching student Admission No: {admission_no}")  # Logging
            dob = request.form.get('dob')

            students_df = pd.read_excel(STUDENTS_FILE)

            # Column name normalization
            students_df.columns = students_df.columns.str.strip()

            # Convert DOB in excel to yyyy-mm-dd
            if 'Date of Birth' in students_df.columns:
                students_df['Date of Birth'] = pd.to_datetime(
                    students_df['Date of Birth'], errors='coerce'
                ).dt.strftime('%Y-%m-%d')

            if 'Admission Number' not in students_df.columns or 'Date of Birth' not in students_df.columns:
                return render_template(
                    'search.html',
                    error="Excel file must have columns: 'Admission Number' and 'Date of Birth'"
                )

            # Student lookup
            student_row = students_df[
                (students_df['Admission Number'].astype(str).str.strip() == admission_no) &
                (students_df['Date of Birth'] == dob)
            ]

            if student_row.empty:
                error = "Student not found or DOB mismatch."
                logger.warning("Student not found or DOB mismatch.")  # Logging warning
                return render_template('search.html', error=error)
            else:
                student_data = student_row.iloc[0].to_dict()
                logger.info("Student found. Rendering index page.")  # Logging success
                return render_template('index.html', student_data=student_data)

        except Exception as e:
            logger.error(f"Error in index route: {str(e)}", exc_info=True)  # Error with traceback
            return f"Error while searching student: {str(e)}"

    # GET request
    logger.info("Rendering search page (GET request)")  # GET log
    return render_template('search.html')


@app.route('/submit', methods=['POST'])
def submit():
    logger.info("Submit route called")  # Logging start
    try:
        admission_no = str(request.form.get('admission_no')).strip()
        logger.info(f"Processing submission for Admission No: {admission_no}")  # Logging
        data = request.form.to_dict()

        # -----------------------
        # HANDLE FILE UPLOAD
        # -----------------------
        file_fields = [
            'ladli_certificate', 'caste_certificate', 'domicile',
            'handicapped_certificate', 'bank_passbook', 'income_certificate'
        ]

        uploaded_files = []
        upload_folder = 'uploads'
        os.makedirs(upload_folder, exist_ok=True)

        for field in file_fields:
            file = request.files.get(field)
            if file and file.filename != '':
                filepath = os.path.join(upload_folder, file.filename)
                logger.info(f"Saving uploaded file: {filepath}")  # Logging file save
                file.save(filepath)
                uploaded_files.append(filepath)

        # -----------------------
        # SEND EMAIL
        # -----------------------
        subject = f"New Submission from Admission No: {admission_no}"
        body = f"Form data submitted:\n\n{data}"
        logger.info("Sending email with attachments")  # Logging
        email_status = send_email_with_files(subject, body, uploaded_files)
        logger.info(f"Email send status: {email_status}")  # Logging result

        # -----------------------
        # APPEND TO EXCEL
        # -----------------------
# -----------------------
# APPEND TO CSV
# -----------------------
        logger.info(f"Appending data to CSV file: {SUBMITTED_FILE}")  # Logging
        import csv

        fieldnames = list(data.keys())
        file_exists = os.path.isfile(SUBMITTED_FILE)

        with open(SUBMITTED_FILE, 'a', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            writer.writerow(data)

        logger.info("CSV file updated successfully")  # Logging



        
        

        # Uploaded temp files ko delete kar dete hain
        for f in uploaded_files:
            if os.path.exists(f):
                os.remove(f)
                logger.info(f"Deleted temp file: {f}")  # Logging file deletion

        if email_status is True:
            logger.info("Form submitted and email sent successfully")  # Log success
            return "Form submitted successfully and email sent!"
        else:
            logger.warning(f"Form submitted but email failed: {email_status}")  # Log warning
            return f"Form submitted but {email_status}"

    except Exception as e:
        logger.error(f"Error in submit route: {str(e)}", exc_info=True)  # Error log with traceback
        return f"Error while submitting form: {str(e)}"


# -----------------------
# SECURE DOWNLOAD ROUTE
# -----------------------
@app.route('/download/<key>', methods=['GET'])
def download(key):
    SECRET_KEY = "shahid-only-download-2025"
    if key != SECRET_KEY:
        return "Unauthorized access", 403

    try:
        return send_file(SUBMITTED_FILE, as_attachment=True)
    except Exception as e:
        return f"Error while downloading file: {str(e)}"


if __name__ == '__main__':
    app.run(debug=True)
