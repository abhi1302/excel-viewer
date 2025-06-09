from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

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
    if str(df.iloc[3, 0]).strip() != "BU PLMN Code":
        messages.append(f"A4 = '{df.iloc[3,0]}' ≠ 'BU PLMN Code'")
    if str(df.iloc[3, 1]).strip() != "TADIG PLMN Code":
        messages.append(f"B4 = '{df.iloc[3,1]}' ≠ 'TADIG PLMN Code'")
    if str(df.iloc[3, 2]).strip() != "Start date":
        messages.append(f"C4 = '{df.iloc[3,2]}' ≠ 'Start date'")
    if str(df.iloc[3, 3]).strip() != "End date":
        messages.append(f"D4 = '{df.iloc[3,3]}' ≠ 'End date'")
    if str(df.iloc[3, 4]).strip() != "Currency":
        messages.append(f"E4 = '{df.iloc[3,4]}' ≠ 'Currency'")
    return messages

@app.route("/", methods=["GET", "POST"])
def index():
    data = None
    headers = []
    errors = []
    row_start = int(request.form.get("start_row", 7))

    if request.method == "POST":
        file = request.files.get("file")
        if file and file.filename.endswith(".xlsx"):
            df_raw = pd.read_excel(file, header=None)
            errors = validate_excel(df_raw)
            if not errors:
                df_data = pd.read_excel(file, header=None, skiprows=row_start - 1)
                df_data.columns = EXPECTED_HEADERS
                data = df_data.fillna("").to_dict(orient="records")
                headers = df_data.columns.tolist()

    return render_template("index.html", data=data, headers=headers, errors=errors, start_row=row_start)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
