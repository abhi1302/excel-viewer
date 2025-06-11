import os
import base64
import io
import json
import logging
import pandas as pd
from flask import Flask, render_template, request, flash, redirect, url_for, session, send_file
from flask_session import Session
from sqlalchemy import create_engine, Column, String, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey")

# Configure Flask-Session
app.config["SESSION_TYPE"] = "filesystem"
app.config["SESSION_FILE_DIR"] = "./.flask_session/"
app.config["SESSION_PERMANENT"] = False
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
app.config["SESSION_COOKIE_NAME"] = "session"  # explicitly set cookie name

# Set session_cookie_name explicitly for Flask-Session
app.session_cookie_name = app.config.get("SESSION_COOKIE_NAME", "session")

Session(app)

# Logging setup
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# Configure SQLAlchemy using Render's external database URL.
PG_DATABASE_URL = os.environ.get("PG_DATABASE_URL", 
    "postgresql://render_postgres_db_7cik_user:123456@dpg-d1479obipnbc73c4hts0-a.oregon-postgres.render.com/render_postgres_db_7cik")
engine = create_engine(PG_DATABASE_URL)
Base = declarative_base()

class Country(Base):
    __tablename__ = 'countries'
    iso_code = Column(String(3), primary_key=True)
    country_name = Column(Text, nullable=False)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# Function to validate the Excel header.
def validate_excel(df):
    """
    Validates the header row (assumed to be row 4, index 3) of the Excel sheet.
    It checks that the expected header values appear in that row.
    (Adjust the expected values as necessary.)
    """
    messages = []
    try:
        # In your original example you expected only 5 columns; now assume the header row actually contains 29 values.
        # Here, we define an expected list (for demo purposes, we assume it as a list with 29 items).
        # You need to adjust these values to match your file's expectations.
        expected_headers = ["Header" + str(i + 1) for i in range(29)]
        # Read row 4 (index 3) to compare:
        actual_headers = [str(x).strip() for x in df.iloc[3, :len(expected_headers)]]
        for i, (expected, actual) in enumerate(zip(expected_headers, actual_headers)):
            if actual != expected:
                messages.append(f"Cell {chr(65+i)}4 = '{actual}' ≠ '{expected}'")
                logger.debug("Validation mismatch at column %s: expected '%s', got '%s'", 
                             chr(65+i), expected, actual)
    except Exception as e:
        err_msg = f"Error during header validation: {e}"
        messages.append(err_msg)
        logger.exception("Exception during Excel header validation")
    return messages

# Function to generate an HTML preview from the uploaded Excel file.
def generate_preview_html():
    """
    Reads the uploaded file from session, loads it via pandas,
    and returns an HTML preview (first 10 rows).
    """
    original_file_b64 = session.get("original_file")
    if not original_file_b64:
        return "<p>No file uploaded yet.</p>"
    try:
        file_bytes = base64.b64decode(original_file_b64)
        stream = io.BytesIO(file_bytes)
        df_preview = pd.read_excel(stream, header=None)
        preview_html = df_preview.head(10).to_html(classes="preview-table", index=False, border=1)
        return preview_html
    except Exception as e:
        logger.exception("Error generating preview")
        return f"<p>Error generating preview: {e}</p>"

@app.route("/", methods=["GET", "POST"])
def index():
    logger.debug("Entered index route with method: %s", request.method)
    # Check if file was uploaded.
    uploaded = "original_file" in session
    data, headers = None, None

    if request.method == "POST":
        step = request.form.get("step")
        logger.debug("Form submitted with step: %s", step)

        if step == "upload":
            file = request.files.get("file")
            if not file:
                flash("No file selected. Please upload a valid Excel file.", "error")
                logger.error("No file selected for upload.")
                return redirect(url_for("index"))

            logger.debug("Received file: %s", file.filename)
            if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
                try:
                    file_bytes = file.read()
                    logger.debug("Read %d bytes from file", len(file_bytes))
                    session["original_file"] = base64.b64encode(file_bytes).decode("utf-8")
                    session["original_filename"] = file.filename
                    flash("File uploaded successfully. Please review the preview and set parameters.", "info")
                    logger.debug("File stored in session successfully.")
                    uploaded = True
                except Exception as e:
                    flash(f"Error while uploading file: {e}", "error")
                    logger.exception("Exception during file upload")
            else:
                flash("Invalid file type. Only .xlsx and .xls files are allowed.", "error")
                logger.error("Invalid file type: %s", file.filename)
            return redirect(url_for("index"))

        elif step == "validate":
            logger.debug("Starting validation process.")
            if not uploaded:
                flash("No file found. Please upload an Excel file first.", "error")
                logger.error("Validation failed: No file in session.")
                return redirect(url_for("index"))

            try:
                file_bytes = base64.b64decode(session["original_file"])
                # Validate header using entire file (or at least row 4)
                stream = io.BytesIO(file_bytes)
                df_raw = pd.read_excel(stream, header=None)
                validation_errors = validate_excel(df_raw)

                if validation_errors:
                    flash("Validation errors detected: " + ", ".join(validation_errors), "error")
                    logger.debug("Validation errors: %s", validation_errors)
                else:
                    flash("Validation successful!", "success")
                    logger.debug("Validation passed without errors.")

                    row_start = int(request.form.get("start_row", 7))
                    # Re-read file bytes to reset the stream.
                    # Here, we no longer limit the number of columns.
                    df_data = pd.read_excel(io.BytesIO(file_bytes), header=None, skiprows=row_start - 1)
                    
                    # Do not slice columns—show all columns (e.g., 29 columns).
                    data = df_data.fillna("").to_dict(orient="records")
                    headers = df_data.columns.tolist()

                    session["validated_data"] = json.dumps(data, default=str)
                    session["validated_headers"] = json.dumps(headers)
                    logger.debug("Validated data stored in session.")

            except Exception as e:
                flash(f"Error processing validation: {e}", "error")
                logger.exception("Exception during validation step")
                return redirect(url_for("index"))

    if "validated_data" in session and "validated_headers" in session:
        data = json.loads(session["validated_data"])
        headers = json.loads(session["validated_headers"])

    preview_html = generate_preview_html() if uploaded else ""
    return render_template("index.html", uploaded=uploaded, preview_html=preview_html, start_row=7, data=data, headers=headers)

@app.route("/download_original")
def download_original():
    """Download the originally uploaded Excel file."""
    original_file_b64 = session.get("original_file")
    original_filename = session.get("original_filename", "uploaded_file.xlsx")
    if not original_file_b64:
        flash("No file available for download.", "error")
        return redirect(url_for("index"))
    try:
        file_bytes = base64.b64decode(original_file_b64)
        return send_file(
            io.BytesIO(file_bytes),
            download_name=original_filename,
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"Error downloading original file: {e}", "error")
        logger.exception("Exception during original file download")
        return redirect(url_for("index"))

@app.route("/download_csv")
def download_csv():
    """Download the uploaded file converted to CSV format."""
    original_file_b64 = session.get("original_file")
    if not original_file_b64:
        flash("No file available for download.", "error")
        return redirect(url_for("index"))
    try:
        file_bytes = base64.b64decode(original_file_b64)
        df = pd.read_excel(io.BytesIO(file_bytes), header=None)
        # Do not slice columns here so that all columns (e.g., 29) are converted.
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        return send_file(
            io.BytesIO(csv_bytes),
            download_name="converted_file.csv",
            as_attachment=True,
            mimetype="text/csv"
        )
    except Exception as e:
        flash(f"Error downloading CSV file: {e}", "error")
        logger.exception("Exception during CSV file download")
        return redirect(url_for("index"))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    logger.debug("Starting Flask server on port %d", port)
    app.run(debug=True, host="0.0.0.0", port=port)
