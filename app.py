
from flask import Flask, render_template_string, send_file
import pandas as pd

app = Flask(__name__)

EXCEL_FILE = "trip_report_detailed.xlsx"

@app.route('/')
def index():
    try:
        df = pd.read_excel(EXCEL_FILE)
        table_html = df.to_html(classes='table table-bordered', index=False)
    except Exception as e:
        table_html = f"<p>خطأ في قراءة التقرير: {str(e)}</p>"

    return render_template_string("""
    <!doctype html>
    <html lang="ar">
    <head>
        <meta charset="utf-8">
        <title>لوحة تقارير مرسانا</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body { padding: 30px; direction: rtl; background-color: #f5f5f5; }
            .container { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px #ccc; }
        </style>
    </head>
    <body>
        <div class="container">
            <h2 class="mb-4">تقرير رحلات مرسانا</h2>
            <a class="btn btn-success mb-3" href="/download">تحميل Excel</a>
            {{ table_html|safe }}
        </div>
    </body>
    </html>
    """, table_html=table_html)

@app.route('/download')
def download():
    return send_file(EXCEL_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
