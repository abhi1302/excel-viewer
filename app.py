from flask import Flask, request, render_template, send_file
import pandas as pd
import io

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    table_html = None
    download_link = None

    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file)

            # Example of formatting/mapping: uppercasing all column names
            df.columns = [col.upper() for col in df.columns]

            # Save processed Excel to memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Processed')
            output.seek(0)

            # Store output in a Flask global for downloading
            global processed_excel
            processed_excel = output

            table_html = df.to_html(classes='table table-striped', index=False)
            download_link = True

    return render_template('index.html', table=table_html, download=download_link)

@app.route('/download')
def download():
    return send_file(processed_excel, download_name="processed_ratecard.xlsx", as_attachment=True)

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
    
