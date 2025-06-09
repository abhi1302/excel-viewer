from flask import Flask, request, render_template, send_file
import pandas as pd
import io
import os

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    table_html = None
    download_link = None

    if request.method == 'POST':
        file = request.files['file']
        start_row = int(request.form.get('start_row', 7))
        if file and file.filename.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file, skiprows=start_row - 1)
            df.columns = [col.upper() for col in df.columns]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Processed')
            output.seek(0)

            global processed_excel
            processed_excel = output

            table_html = df.to_html(classes='table table-striped', index=False)
            download_link = True

    return render_template('index.html', table=table_html, download=download_link)

@app.route('/download')
def download():
    return send_file(processed_excel, download_name="processed_ratecard.xlsx", as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
