import os
import io
import json
import base64
import logging
import pandas as pd
from flask import Flask, render_template, request, flash, redirect, url_for, Response, session
from flask_session import Session  # Import Flask-Session

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey")

# Monkey-patch to add session_cookie_name for compatibility
if not hasattr(app, 'session_cookie_name'):
    app.session_cookie_name = app.config.get("SESSION_COOKIE_NAME", "session")

# Configure Flask-Session for server-side sessions
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
    data = None
    headers = []
    errors = []
    row_start = int(request.form.get("start_row", 7))
    logger.debug("Row start set to: %d", row_start)
    
    # Default row count value to pass to template
    row_count = None

    if request.method == "POST":
        file = request.files.get("file")
        if file and (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
            logger.debug("Received file: %s", file.filename)
            try:
                file_bytes = file.read()
                logger.debug("File size: %d bytes", len(file_bytes))
                session['original_file'] = base64.b64encode(file_bytes).decode('utf-8')
                session['original_filename'] = file.filename

                stream_validation = io.BytesIO(file_bytes)
                df_raw = pd.read_excel(stream_validation, header=None)
                logger.debug("Excel file read for header validation successfully")
                errors = validate_excel(df_raw)

                if errors:
                    flash("Excel validation errors encountered.", "error")
                    logger.debug("Excel validation errors: %s", errors)
                else:
                    stream_data = io.BytesIO(file_bytes)
                    df_data = pd.read_excel(stream_data, header=None, skiprows=row_start - 1)
                    logger.debug("Excel file read for data processing successfully")
                    df_data.columns = EXPECTED_HEADERS
                    data = df_data.fillna("").to_dict(orient="records")
                    headers = df_data.columns.tolist()

                    # Compute the row count for the data (total number of data rows)
                    row_count = df_data.shape[0]
                    logger.debug("Row count computed: %d", row_count)

                    session['data'] = json.dumps(data, default=str)
                    session['headers'] = json.dumps(headers)
                    logger.debug("Data, headers, and row count stored in session")
            except Exception as e:
                error_msg = f"Failed to process excel file: {e}"
                errors.append(error_msg)
                flash(error_msg, "error")
                logger.exception("Exception during file processing")
        else:
            error_msg = "Please upload a valid Excel file (.xlsx or .xls)"
            errors.append(error_msg)
            flash(error_msg, "error")
            logger.debug("Invalid file provided or file type unsupported")
    else:
        logger.debug("GET method for index route")

    # Pass row_count to template (it may be None if not set)
    return render_template("index.html", data=data, headers=headers, errors=errors, start_row=row_start, row_count=row_count)

@app.route("/download_original")
def download_original():
    logger.debug("Download original file requested")
    original_file_b64 = session.get('original_file')
    original_filename = session.get('original_filename', "original_file.xlsx")
    if not original_file_b64:
        flash("No original file available for download. Please upload an Excel file first.", "error")
        logger.debug("No original file in session")
        return redirect(url_for("index"))
    
    file_bytes = base64.b64decode(original_file_b64)
    if original_filename.endswith(".xls"):
        mimetype = "application/vnd.ms-excel"
    else:
        mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
    logger.debug("Returning original file: %s", original_filename)
    return Response(
        file_bytes,
        mimetype=mimetype,
        headers={"Content-Disposition": f"attachment; filename={original_filename}"}
    )

@app.route("/download_csv")
def download_csv():
    logger.debug("Download CSV requested")
    data_json = session.get('data')
    headers_json = session.get('headers')
    if not data_json or not headers_json:
        flash("No data available for download. Please upload an Excel file first.", "error")
        logger.debug("No processed data in session")
        return redirect(url_for("index"))
    
    data = json.loads(data_json)
    headers = json.loads(headers_json)
    df = pd.DataFrame(data, columns=headers)

    logger.debug("Converting processed data to CSV")
    csv_buffer = io.StringIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    
    return Response(
        csv_buffer,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=excel_data.csv"}
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    logger.debug("Starting Flask server on port %d", port)
    app.run(debug=True, host="0.0.0.0", port=port)
