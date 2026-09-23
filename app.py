from flask import Flask, render_template_string, send_file
import pandas as pd
import os

app = Flask(__name__)

EXCEL_FILE = "naukri_jobs.xlsx"

HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Naukri Job Scraper</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f4f6f8;
            margin: 0;
            padding: 20px;
        }

        .container {
            max-width: 1400px;
            margin: auto;
        }

        h1 {
            text-align: center;
        }

        .info {
            text-align: center;
            color: #555;
            margin-bottom: 20px;
        }

        .download {
            display: inline-block;
            padding: 10px 18px;
            background: #1976d2;
            color: white;
            text-decoration: none;
            border-radius: 6px;
            margin-bottom: 20px;
        }

        .table-container {
            overflow-x: auto;
            background: white;
            padding: 15px;
            border-radius: 10px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        th, td {
            border: 1px solid #ddd;
            padding: 10px;
            text-align: left;
        }

        th {
            background: #1976d2;
            color: white;
        }

        tr:nth-child(even) {
            background: #f2f2f2;
        }

        a {
            color: #1565c0;
        }
    </style>
</head>

<body>

<div class="container">

    <h1>🔎 Naukri Job Scraper</h1>

    <div class="info">
        Python Developer Jobs — Scraped using Python & Playwright
    </div>

    <a class="download" href="/download">
        Download Excel
    </a>

    <div class="table-container">
        {{ table|safe }}
    </div>

</div>

</body>
</html>
"""


@app.route("/")
def home():

    if not os.path.exists(EXCEL_FILE):
        return "Excel file not found."

    df = pd.read_excel(EXCEL_FILE)

    if "Job URL" in df.columns:
        df["Job URL"] = df["Job URL"].apply(
            lambda x: f'<a href="{x}" target="_blank">View Job</a>'
        )

    table = df.to_html(
        index=False,
        escape=False
    )

    return render_template_string(HTML, table=table)


@app.route("/download")
def download():

    return send_file(
        EXCEL_FILE,
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)