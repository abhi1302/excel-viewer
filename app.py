from flask import Flask, render_template, request, flash, redirect, url_for, Response, session
import pandas as pd
import os
import io
import json

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

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
    except Exception as e:
        messages.append(f"Error during header validation: {e}")
    return messages

@app.route("/", methods=["GET", "POST"])
def index():
    data = None
    headers = []
    errors = []
    row_start = int(request.form.get("start_row", 7))

    if request.method == "POST":
        file = request.files.get("file")
        if file and (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
            try:
                # Read file for header validation
                df_raw = pd.read_excel(file, header=None)
                errors = validate_excel(df_raw)

                if errors:
                    flash("Excel validation errors encountered.", "error")
                else:
                    file.seek(0)
                    df_data = pd.read_excel(file, header=None, skiprows=row_start - 1)
                    df_data.columns = EXPECTED_HEADERS
                    data = df_data.fillna("").to_dict(orient="records")
                    headers = df_data.columns.tolist()
                    
                    # Use default=str to convert Pandas Timestamps to strings
                    session['data'] = json.dumps(data, default=str)
                    session['headers'] = json.dumps(headers)
            except Exception as e:
                errors.append(f"Failed to process excel file: {e}")
                flash(f"Failed to process excel file: {e}", "error")
        else:
            errors.append("Please upload a valid Excel file (.xlsx or .xls)")
            flash("Please upload a valid Excel file (.xlsx or .xls)", "error")

    return render_template("index.html", data=data, headers=headers, errors=errors, start_row=row_start)

@app.route("/download")
def download():
    # Retrieve data from session
    data_json = session.get('data')
    headers_json = session.get('headers')
    if not data_json or not headers_json:
        flash("No data available for download. Please upload an Excel file first.", "error")
        return redirect(url_for("index"))
    
    data = json.loads(data_json)
    headers = json.loads(headers_json)
    
    # Convert data into a DataFrame
    df = pd.DataFrame(data, columns=headers)
    
    # Save DataFrame to CSV in a memory buffer
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
    app.run(debug=True, host="0.0.0.0", port=port)
