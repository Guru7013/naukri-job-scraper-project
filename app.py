from flask import Flask, render_template_string, send_file, request
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = "naukri_jobs.xlsx"


HTML = """
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

    <title>Naukri Job Scraper</title>

    <style>

        * {
            box-sizing: border-box;
        }

        body {
            margin: 0;
            font-family: Arial, Helvetica, sans-serif;
            background: #f4f7fb;
            color: #1f2937;
            transition: 0.3s;
        }

        body.dark {
            background: #111827;
            color: #f3f4f6;
        }

        .header {
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            color: white;
            padding: 38px 20px;
            text-align: center;
        }

        .header h1 {
            margin: 0;
            font-size: 36px;
        }

        .header p {
            margin: 10px 0 0;
            font-size: 16px;
            opacity: 0.92;
        }

        .top-buttons {
            margin-top: 18px;
        }

        .dark-btn {
            background: white;
            color: #1d4ed8;
            border: none;
            padding: 9px 15px;
            border-radius: 20px;
            cursor: pointer;
            font-weight: bold;
        }

        .container {
            max-width: 1450px;
            margin: 28px auto;
            padding: 0 20px;
        }

        .stats {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 18px;
            margin-bottom: 25px;
        }

        .card {
            background: white;
            padding: 24px;
            border-radius: 14px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.08);
            text-align: center;
        }

        body.dark .card,
        body.dark .toolbar,
        body.dark .table-container {
            background: #1f2937;
            color: #f3f4f6;
        }

        .card h2 {
            margin: 0;
            color: #2563eb;
            font-size: 32px;
        }

        .card p {
            margin: 8px 0 0;
            color: #6b7280;
        }

        body.dark .card p {
            color: #d1d5db;
        }

        .toolbar {
            background: white;
            padding: 20px;
            border-radius: 14px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.08);
            margin-bottom: 20px;
        }

        .toolbar form {
            display: grid;
            grid-template-columns: 2fr 1fr 1fr 1fr auto auto;
            gap: 12px;
        }

        input,
        select {
            padding: 12px;
            border: 1px solid #d1d5db;
            border-radius: 8px;
            font-size: 14px;
            outline: none;
            background: white;
            color: #111827;
        }

        body.dark input,
        body.dark select {
            background: #374151;
            color: white;
            border-color: #4b5563;
        }

        input:focus,
        select:focus {
            border-color: #2563eb;
        }

        button,
        .clear-btn,
        .download {
            border: none;
            padding: 12px 18px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: bold;
            text-decoration: none;
            text-align: center;
        }

        .search-btn {
            background: #2563eb;
            color: white;
        }

        .clear-btn {
            background: #e5e7eb;
            color: #111827;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .download {
            display: inline-block;
            background: #16a34a;
            color: white;
            margin-bottom: 15px;
        }

        .info-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            gap: 10px;
            flex-wrap: wrap;
        }

        .result {
            color: #4b5563;
            font-weight: bold;
        }

        body.dark .result {
            color: #d1d5db;
        }

        .updated {
            font-size: 13px;
            color: #6b7280;
        }

        body.dark .updated {
            color: #d1d5db;
        }

        .table-container {
            background: white;
            border-radius: 14px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.08);
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 1100px;
        }

        th {
            background: #2563eb;
            color: white;
            padding: 15px;
            text-align: left;
            font-size: 14px;
            white-space: nowrap;
        }

        td {
            padding: 14px;
            border-bottom: 1px solid #e5e7eb;
            font-size: 14px;
            vertical-align: top;
        }

        body.dark td {
            border-bottom-color: #374151;
        }

        tr:hover {
            background: #f8fafc;
        }

        body.dark tr:hover {
            background: #374151;
        }

        .job-title {
            font-weight: bold;
        }

        .skills {
            display: flex;
            flex-wrap: wrap;
            gap: 5px;
        }

        .skill {
            display: inline-block;
            background: #dbeafe;
            color: #1d4ed8;
            padding: 5px 8px;
            border-radius: 15px;
            font-size: 11px;
            font-weight: bold;
        }

        body.dark .skill {
            background: #1e3a8a;
            color: #dbeafe;
        }

        .job-link {
            color: #2563eb;
            font-weight: bold;
            text-decoration: none;
        }

        .job-link:hover {
            text-decoration: underline;
        }

        .pagination {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 8px;
            margin: 25px 0;
            flex-wrap: wrap;
        }

        .page-btn {
            background: white;
            color: #2563eb;
            border: 1px solid #d1d5db;
            padding: 9px 13px;
            border-radius: 7px;
            text-decoration: none;
            font-weight: bold;
        }

        body.dark .page-btn {
            background: #1f2937;
            border-color: #4b5563;
            color: #dbeafe;
        }

        .page-btn.active {
            background: #2563eb;
            color: white;
            border-color: #2563eb;
        }

        .footer {
            text-align: center;
            padding: 30px;
            color: #6b7280;
            font-size: 13px;
        }

        body.dark .footer {
            color: #d1d5db;
        }

        @media (max-width: 900px) {

            .stats {
                grid-template-columns: 1fr;
            }

            .toolbar form {
                grid-template-columns: 1fr;
            }

            .header h1 {
                font-size: 28px;
            }
        }

    </style>
</head>


<body>

<div class="header">

    <h1>🔎 Naukri Job Scraper</h1>

    <p>
        Python Developer Jobs — Scraped using Python & Playwright
    </p>

    <div class="top-buttons">

        <button class="dark-btn" onclick="toggleDarkMode()">
            🌙 Dark Mode
        </button>

    </div>

</div>


<div class="container">


    <!-- Statistics -->

    <div class="stats">

        <div class="card">

            <h2>{{ total_jobs }}</h2>

            <p>📋 Total Jobs</p>

        </div>


        <div class="card">

            <h2>{{ total_companies }}</h2>

            <p>🏢 Companies</p>

        </div>


        <div class="card">

            <h2>{{ total_locations }}</h2>

            <p>📍 Locations</p>

        </div>

    </div>


    <!-- Search and filters -->

    <div class="toolbar">

        <form method="GET">

            <input
                type="text"
                name="search"
                placeholder="🔍 Search job title or company..."
                value="{{ search }}"
            >


            <select name="location">

                <option value="">
                    All Locations
                </option>

                {% for loc in locations %}

                <option
                    value="{{ loc }}"
                    {% if loc == selected_location %}selected{% endif %}
                >
                    {{ loc }}
                </option>

                {% endfor %}

            </select>


            <select name="experience">

                <option value="">
                    All Experience
                </option>

                {% for exp in experiences %}

                <option
                    value="{{ exp }}"
                    {% if exp == selected_experience %}selected{% endif %}
                >
                    {{ exp }}
                </option>

                {% endfor %}

            </select>


            <select name="sort">

                <option value="title"
                    {% if sort_by == "title" %}selected{% endif %}>
                    Sort: Job Title
                </option>

                <option value="company"
                    {% if sort_by == "company" %}selected{% endif %}>
                    Sort: Company
                </option>

                <option value="location"
                    {% if sort_by == "location" %}selected{% endif %}>
                    Sort: Location
                </option>

                <option value="experience"
                    {% if sort_by == "experience" %}selected{% endif %}>
                    Sort: Experience
                </option>

            </select>


            <button
                class="search-btn"
                type="submit"
            >
                🔍 Search
            </button>


            <a
                class="clear-btn"
                href="/"
            >
                Clear
            </a>

        </form>

    </div>


    <!-- Download -->

    <a
        class="download"
        href="/download"
    >
        📥 Download Excel
    </a>


    <!-- Information -->

    <div class="info-row">

        <div class="result">

            Showing {{ displayed_jobs }} job(s)

        </div>


        <div class="updated">

            🕐 Last updated: {{ last_updated }}

        </div>

    </div>


    <!-- Table -->

    <div class="table-container">

        <table>

            <thead>

                <tr>

                    <th>Job Title</th>

                    <th>Company</th>

                    <th>Location</th>

                    <th>Experience</th>

                    <th>Skills</th>

                    <th>Posted Date</th>

                    <th>Job URL</th>

                </tr>

            </thead>


            <tbody>

                {% for job in jobs %}

                <tr>

                    <td class="job-title">

                        {{ job["Job Title"] }}

                    </td>


                    <td>

                        {{ job["Company"] }}

                    </td>


                    <td>

                        {{ job["Location"] }}

                    </td>


                    <td>

                        {{ job["Experience"] }}

                    </td>


                    <td>

                        <div class="skills">

                            {% set skill_text = job["Skills"] | string %}

                            {% for skill in skill_text.split(",") %}

                                {% if skill.strip() %}

                                <span class="skill">

                                    {{ skill.strip() }}

                                </span>

                                {% endif %}

                            {% endfor %}

                        </div>

                    </td>


                    <td>

                        {{ job["Posted Date"] }}

                    </td>


                    <td>

                        {% if job["Job URL"] %}

                        <a
                            class="job-link"
                            href="{{ job["Job URL"] }}"
                            target="_blank"
                        >
                            View Job
                        </a>

                        {% else %}

                        -

                        {% endif %}

                    </td>

                </tr>

                {% endfor %}

            </tbody>

        </table>

    </div>


    <!-- Pagination -->

    {% if total_pages > 1 %}

    <div class="pagination">

        {% if page > 1 %}

        <a
            class="page-btn"
            href="{{ url_for('home',
                page=page-1,
                search=search,
                location=selected_location,
                experience=selected_experience,
                sort=sort_by) }}"
        >
            ← Previous
        </a>

        {% endif %}


        {% for p in range(1, total_pages + 1) %}

            {% if p == page %}

            <a
                class="page-btn active"
                href="#"
            >
                {{ p }}
            </a>

            {% else %}

            <a
                class="page-btn"
                href="{{ url_for('home',
                    page=p,
                    search=search,
                    location=selected_location,
                    experience=selected_experience,
                    sort=sort_by) }}"
            >
                {{ p }}
            </a>

            {% endif %}

        {% endfor %}


        {% if page < total_pages %}

        <a
            class="page-btn"
            href="{{ url_for('home',
                page=page+1,
                search=search,
                location=selected_location,
                experience=selected_experience,
                sort=sort_by) }}"
        >
            Next →
        </a>

        {% endif %}

    </div>

    {% endif %}


</div>


<div class="footer">

    Naukri Job Scraper | Python + Playwright + Pandas

</div>


<script>

function toggleDarkMode() {

    document.body.classList.toggle("dark");

    if (document.body.classList.contains("dark")) {

        localStorage.setItem("darkMode", "enabled");

    } else {

        localStorage.setItem("darkMode", "disabled");

    }

}


if (localStorage.getItem("darkMode") === "enabled") {

    document.body.classList.add("dark");

}

</script>


</body>

</html>
"""


@app.route("/")
def home():

    if not os.path.exists(EXCEL_FILE):

        return "Excel file not found."


    try:

        df = pd.read_excel(EXCEL_FILE)

        required_columns = [
            "Job Title",
            "Company",
            "Location",
            "Experience",
            "Skills",
            "Posted Date",
            "Job URL"
        ]


        for column in required_columns:

            if column not in df.columns:

                df[column] = ""


        df = df.fillna("")


        # -------------------------
        # Statistics
        # -------------------------

        total_jobs = len(df)

        total_companies = (
            df["Company"]
            .astype(str)
            .replace("", pd.NA)
            .nunique()
        )

        total_locations = (
            df["Location"]
            .astype(str)
            .replace("", pd.NA)
            .nunique()
        )


        # -------------------------
        # Filter values
        # -------------------------

        locations = sorted(
            [
                str(x)
                for x in df["Location"].unique()
                if str(x).strip()
            ]
        )


        experiences = sorted(
            [
                str(x)
                for x in df["Experience"].unique()
                if str(x).strip()
            ]
        )


        # -------------------------
        # Request values
        # -------------------------

        search = request.args.get(
            "search",
            ""
        ).strip()


        selected_location = request.args.get(
            "location",
            ""
        ).strip()


        selected_experience = request.args.get(
            "experience",
            ""
        ).strip()


        sort_by = request.args.get(
            "sort",
            "title"
        ).strip()


        try:

            page = int(
                request.args.get(
                    "page",
                    1
                )
            )

        except ValueError:

            page = 1


        if page < 1:

            page = 1


        # -------------------------
        # Filtering
        # -------------------------

        filtered_df = df.copy()


        if search:

            search_lower = search.lower()

            title_match = (
                filtered_df["Job Title"]
                .astype(str)
                .str.lower()
                .str.contains(
                    search_lower,
                    na=False
                )
            )


            company_match = (
                filtered_df["Company"]
                .astype(str)
                .str.lower()
                .str.contains(
                    search_lower,
                    na=False
                )
            )


            filtered_df = filtered_df[
                title_match | company_match
            ]


        if selected_location:

            filtered_df = filtered_df[
                filtered_df["Location"]
                .astype(str)
                == selected_location
            ]


        if selected_experience:

            filtered_df = filtered_df[
                filtered_df["Experience"]
                .astype(str)
                == selected_experience
            ]


        # -------------------------
        # Sorting
        # -------------------------

        sort_columns = {

            "title": "Job Title",

            "company": "Company",

            "location": "Location",

            "experience": "Experience"

        }


        sort_column = sort_columns.get(
            sort_by,
            "Job Title"
        )


        filtered_df = filtered_df.sort_values(
            by=sort_column,
            key=lambda x: x.astype(str).str.lower()
        )


        # -------------------------
        # Pagination
        # -------------------------

        per_page = 10

        total_filtered_jobs = len(filtered_df)

        total_pages = max(
            1,
            (total_filtered_jobs + per_page - 1)
            // per_page
        )


        if page > total_pages:

            page = total_pages


        start = (page - 1) * per_page

        end = start + per_page


        page_df = filtered_df.iloc[
            start:end
        ]


        jobs = page_df.to_dict(
            "records"
        )


        # -------------------------
        # Last updated
        # -------------------------

        modified_time = os.path.getmtime(
            EXCEL_FILE
        )


        last_updated = datetime.fromtimestamp(
            modified_time
        ).strftime(
            "%d %b %Y, %I:%M %p"
        )


        return render_template_string(

            HTML,

            jobs=jobs,

            total_jobs=total_jobs,

            total_companies=total_companies,

            total_locations=total_locations,

            locations=locations,

            experiences=experiences,

            search=search,

            selected_location=selected_location,

            selected_experience=selected_experience,

            sort_by=sort_by,

            displayed_jobs=total_filtered_jobs,

            page=page,

            total_pages=total_pages,

            last_updated=last_updated

        )


    except Exception as e:

        return f"Error loading jobs: {e}"


@app.route("/download")
def download():

    if not os.path.exists(EXCEL_FILE):

        return "Excel file not found."


    return send_file(

        EXCEL_FILE,

        as_attachment=True,

        download_name="naukri_jobs.xlsx"

    )


if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )


    app.run(

        host="0.0.0.0",

        port=port,

        debug=False

    )