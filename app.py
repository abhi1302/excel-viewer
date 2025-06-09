from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

HEADERS = [
    "BU PLMN Code", "TADIG PLMN Code", "Start date", "End date", "Currency",
    "MOC Local Call Rate/Value", "Charging interval",
    "MOC Call Back Home Rate/Value", "Charging interval",
    "MOC Rest of the world Rate/Value", "Charging interval",
    "MOC Premium numbers Rate/Value", "Charging interval",
    "MOC Special numbers Rate/Value", "Charging interval",
    "MOC Satellite Rate/Value", "Charging interval",
    "MTC Call Rate/Value", "Charging interval",
    "MO-SMS Rate/Value",
    "GPRS Rate MB Rate/Value", "GPRS Rate per KB Rate/Value", "Charging interval",
    "VoLTE Rate MB Rate/Value", "Charging interval",
    "Tax applicable Yes/No", "Tax applicable Tax Value",
    "Tax included in the rate Yes/No",
    "Bearer Service included in Special IOT Yes/No"
]

@app.route("/", methods=["GET", "POST"])
def index():
    error_messages = []
    data = None
    headers = []

    if request.method == "POST":
        file = request.files["file"]
        start_row = int(request.form.get("start_row", 7))

        if file and file.filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file, header=None)
            try:
                for i, expected in enumerate(HEADERS):
                    row = start_row - 1  # zero-indexed
                    col = i
                    actual = str(df.iloc[row, col]).strip() if not pd.isna(df.iloc[row, col]) else ""
                    if actual != expected:
                        error_messages.append(f"Column {i+1}: Expected '{expected}' but found '{actual}'")

                if not error_messages:
                    df_data = df.iloc[start_row:, :len(HEADERS)]
                    df_data.columns = HEADERS
                    data = df_data.to_html(index=False)
                    return render_template("index.html", data=data, headers=headers)

            except Exception as e:
                error_messages.append(f"Unexpected error: {str(e)}")

    return render_template("index.html", data=None, errors=error_messages)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
