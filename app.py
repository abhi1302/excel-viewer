import os
import io
import json
import base64
import logging
import pandas as pd
from flask import Flask, render_template, request, flash, redirect, url_for, Response, session
from flask_session import Session  # Flask-Session for server-side sessions

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey")

# Monkey-patch for compatibility with Flask-Session (Flask 2.3+)
if not hasattr(app, 'session_cookie_name'):
    app.session_cookie_name = app.config.get("SESSION_COOKIE_NAME", "session")

# Configure Flask-Session to use the filesystem
app.config["SESSION_TYPE"] = "filesystem"
app.config["SESSION_FILE_DIR"] = "./.flask_session/"
app.config["SESSION_PERMANENT"] = False
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
Session(app)

# Set up logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# Expected column headers for the Excel data
EXPECTED_HEADERS = [
    "BU PLMN Code",
    "TADIG PLMN Code",
    "Start date",
    "End date",
    "Currency",
    "MOC Local Call Rate/Value",
    "Charging interval",
    "MOC Call Back Home Rate/Value",
    "Charging interval",
    "MOC Rest of the world Rate/Value",
    "Charging interval",
    "MOC Premium numbers Rate/Value",
    "Charging interval",
    "MOC Special numbers Rate/Value",
    "Charging interval",
    "MOC Satellite Rate/Value",
    "Charging interval",
    "MTC Call Rate/Value",
    "Charging interval",
    "MO-SMS Rate/Value",
    "GPRS Rate MB Rate/Value",
    "GPRS Rate per KB Rate/Value",
    "Charging interval",
    "VoLTE Rate MB Rate/Value",
    "Charging interval",
    "Tax applicable Yes/No",
    "Tax applicable Tax Value",
    "Tax included in the rate Yes/No",
    "Bearer Service included in Special IOT Yes/No"
]

def validate_excel(df):
    """Validates the header row (row index 3 or fourth row) of the Excel file."""
    messages = []
    try:
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
        error_msg = f"Error during header validation: {e}"
        messages.append(error_msg)
        logger.exception("Exception during Excel header validation")
    return messages

@app.route("/", methods=["GET", "POST"])
def index():
    logger.debug("Loading index route")
    # If the request is POST, determine which step we're processing
    if request.method == "POST":
        step = request.form.get("step")
        # Step One: file upload
        if step == "upload":
            file = request.files.get("file")
            if file and (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
                logger.debug("Received file: %s", file.filename)
                try:
                    file_bytes = file.read()
                    logger.debug("File size: %d bytes", len(file_bytes))
                    # Store file (Base64 encoded) in session
                    session["original_file"] = base64.b64encode(file_bytes).decode("utf-8")
                    session["original_filename"] = file.filename
                    flash("File uploaded successfully. Now set your parameters and click 'Validate Ratesheet'.", "info")
                except Exception as e:
                    flash(f"Error while uploading file: {e}", "error")
                    logger.exception("Exception during file upload")
            else:
                flash("Please upload a valid Excel file (.xlsx or .xls)", "error")
            # Redirect so that the GET method can show the next step form
            return redirect(url_for("index"))
        
        # Step Two: process file with provided parameters and validate
        elif step == "validate":
            start_row = int(request.form.get("start_row", 7))
            logger.debug("Parameter received: start_row = %d", start_row)
            data = None
            headers = None
            row_count = None

            original_file_b64 = session.get("original_file")
            if not original_file_b64:
                flash("No file uploaded. Please upload an Excel file first.", "error")
                return redirect(url_for("index"))

            try:
                file_bytes = base64.b64decode(original_file_b64)
                # Use a stream to perform header validation
                stream_validation = io.BytesIO(file_bytes)
                df_raw = pd.read_excel(stream_validation, header=None)
                errors = validate_excel(df_raw)
                if errors:
                    flash("Validation errors: " + ", ".join(errors), "error")
                    logger.debug("Validation errors: %s", errors)
                else:
                    # Process file using provided start_row (data starts from that row)
                    stream_data = io.BytesIO(file_bytes)
                    df_data = pd.read_excel(stream_data, header=None, skiprows=start_row - 1)
                    logger.debug("Excel file read for data processing successfully")
                    df_data.columns = EXPECTED_HEADERS
                    data = df_data.fillna("").to_dict(orient="records")
                    headers = df_data.columns.tolist()
                    row_count = df_data.shape[0]
                    # Optionally, store processed data in session for further use (like download)
                    session["data"] = json.dumps(data, default=str)
                    session["headers"] = json.dumps(headers)
                    flash(f"File processed successfully. Total rows: {row_count}", "success")
                    logger.debug("Data processed successfully with row count: %d", row_count)
            except Exception as e:
                flash(f"Failed to process Excel file: {e}", "error")
                logger.exception("Exception during file processing in step 2")
            # Render the page in step two, displaying the results (if available)
            return render_template("index.html", 
                                   uploaded=True, start_row=start_row, 
                                   data=data, headers=headers, row_count=row_count)
    
    # GET request: decide which form to show based on whether a file was already uploaded
    uploaded = "original_file" in session
    return render_template("index.html", uploaded=uploaded)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    logger.debug("Starting Flask server on port %d", port)
    app.run(debug=True, host="0.0.0.0", port=port)
