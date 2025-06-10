import os
import base64
import io
import json
import logging
import pandas as pd
from flask import Flask, render_template, request, flash, redirect, url_for, Response, session
from flask_session import Session
from sqlalchemy import create_engine, Column, String, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey")

# Monkey-patch for Flask-Session in Flask 2.3+
if not hasattr(app, 'session_cookie_name'):
    app.session_cookie_name = app.config.get("SESSION_COOKIE_NAME", "session")

app.config["SESSION_TYPE"] = "filesystem"
app.config["SESSION_FILE_DIR"] = "./.flask_session/"
app.config["SESSION_PERMANENT"] = False
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
Session(app)

# Set up logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# Configure SQLAlchemy using Render's external database URL.
PG_DATABASE_URL = os.environ.get("PG_DATABASE_URL",
    "postgresql://render_postgres_db_7cik_user:123456@dpg-d1479obipnbc73c4hts0-a.oregon-postgres.render.com/render_postgres_db_7cik"
)
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
    Validates the header row (assumed to be row 4, i.e. index 3) of the Excel sheet.
    Checks that the first five cells contain the expected values.
    """
    messages = []
    try:
        # Expected headers for columns A to E (change these as needed)
        validation_checks = [
            ("BU PLMN Code", 0),
            ("TADIG PLMN Code", 1),
            ("Start date", 2),
            ("End date", 3),
            ("Currency", 4)
        ]
        for expected, col_index in validation_checks:
            actual = str(df.iloc[3, col_index]).strip()
            if actual != expected:
                messages.append(f"Cell {chr(65+col_index)}4 = '{actual}' ≠ '{expected}'")
                logger.debug("Validation mismatch at column %s: expected '%s', got '%s'",
                             chr(65+col_index), expected, actual)
    except Exception as e:
        err_msg = f"Error during header validation: {e}"
        messages.append(err_msg)
        logger.exception("Exception during Excel header validation")
    return messages

# A helper function for country lookup from the database.
def get_country_from_tadig(tadig_code):
    lookup_code = tadig_code[:3].upper()
    db = SessionLocal()
    try:
        country_obj = db.query(Country).filter(Country.iso_code == lookup_code).first()
        if country_obj:
            return country_obj.country_name
        else:
            return "Unknown"
    except Exception as e:
        logger.exception("Error looking up country for %s", lookup_code)
        return "Error"
    finally:
        db.close()

# Function to generate an HTML preview from the uploaded Excel file.
def generate_preview_html():
    original_file_b64 = session.get("original_file")
    if not original_file_b64:
        return "<p>No file uploaded yet.</p>"
    try:
        file_bytes = base64.b64decode(original_file_b64)
        stream = io.BytesIO(file_bytes)
        df_preview = pd.read_excel(stream, header=None)
        # Generate preview: First 10 rows as an HTML table.
        preview_html = df_preview.head(10).to_html(classes="preview-table", index=False, border=1)
        return preview_html
    except Exception as e:
        logger.exception("Error generating preview")
        return f"<p>Error generating preview: {e}</p>"

# Excel processing route that integrates the country lookup (if needed in the future).
@app.route("/process_excel", methods=["POST"])
def process_excel():
    file = request.files.get("file")
    if file and file.filename.endswith((".xlsx", ".xls")):
        try:
            logger.debug("Processing uploaded file: %s", file.filename)
            file_bytes = file.read()
            df = pd.read_excel(io.BytesIO(file_bytes), header=0)
            # Check for the expected column and append 'Country' via lookup.
            if "TADIG PLMN Code" in df.columns:
                df["Country"] = df["TADIG PLMN Code"].apply(get_country_from_tadig)
            else:
                flash("Error: Column 'TADIG PLMN Code' not found in the uploaded file.", "error")
                logger.error("Column 'TADIG PLMN Code' missing in uploaded Excel file.")
                return redirect(url_for("index"))
            table_html = df.to_html(classes="data-table", index=False)
            return render_template("results.html", table_html=table_html)
        except Exception as e:
            logger.exception("Error processing Excel file")
            flash(f"Error processing Excel file: {e}", "error")
    else:
        flash("Please upload a valid Excel file.", "error")
    return redirect(url_for("index"))

@app.route("/", methods=["GET", "POST"])
def index():
    logger.debug("Entered index route with method: %s", request.method)
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
                except Exception as e:
                    flash(f"Error while uploading file: {e}", "error")
                    logger.exception("Exception during file upload")
            else:
                flash("Invalid file type. Only .xlsx and .xls files are allowed.", "error")
                logger.error("Invalid file type: %s", file.filename)
            return redirect(url_for("index"))

        elif step == "validate":
            logger.debug("Starting validation process.")
            original_file_b64 = session.get("original_file")
            if not original_file_b64:
                flash("No file found. Please upload an Excel file first.", "error")
                logger.error("Validation failed: No file in session.")
                return redirect(url_for("index"))
            try:
                file_bytes = base64.b64decode(original_file_b64)
                stream = io.BytesIO(file_bytes)
                df_raw = pd.read_excel(stream, header=None)
                validation_errors = validate_excel(df_raw)
                if validation_errors:
                    flash("Validation errors detected: " + ", ".join(validation_errors), "error")
                    logger.debug("Validation errors: %s", validation_errors)
                else:
                    flash("Validation successful!", "success")
                    logger.debug("Validation passed without errors.")
            except Exception as e:
                flash(f"Error processing validation: {e}", "error")
                logger.exception("Exception during validation step")
            preview_html = generate_preview_html()
            return render_template("index.html", uploaded=True, preview_html=preview_html, start_row=7)

    # For GET requests or after redirects:
    if "original_file" in session:
        preview_html = generate_preview_html()
        logger.debug("Rendering GET with preview_html")
        return render_template("index.html", uploaded=True, preview_html=preview_html, start_row=7)
    else:
        logger.debug("Rendering GET without file uploaded")
        return render_template("index.html", uploaded=False)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    logger.debug("Starting Flask server on port %d", port)
    app.run(debug=True, host="0.0.0.0", port=port)