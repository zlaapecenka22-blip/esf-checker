from flask import (
    Flask,
    render_template,
    request,
    send_file
)

import os

from checker import compare_files
from report_generator import create_report

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
REPORT_FOLDER = "reports"

os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)

os.makedirs(
    REPORT_FOLDER,
    exist_ok=True
)


@app.route("/")
def home():

    return render_template(
        "index.html"
    )


@app.route(
    "/check",
    methods=["POST"]
)
def check():

    file_1c = request.files["file_1c"]
    file_esf = request.files["file_esf"]

    path_1c = os.path.join(
        UPLOAD_FOLDER,
        file_1c.filename
    )

    path_esf = os.path.join(
        UPLOAD_FOLDER,
        file_esf.filename
    )

    file_1c.save(path_1c)
    file_esf.save(path_esf)

    results = compare_files(
        path_1c,
        path_esf
    )

    create_report(results)

    total = len(results)

    ok_count = len(
        [
            r for r in results
            if r["Статус"] == "OK"
        ]
    )

    error_count = total - ok_count

    match_percent = round(
        (ok_count / total) * 100,
        2
    ) if total else 0

    matches = [
        r for r in results
        if r["Тип"] == "Совпадение"
    ]

    differences = [
        r for r in results
        if r["Тип"] == "Расхождение"
    ]

    duplicates = [
        r for r in results
        if r["Тип"] == "Дубликат"
    ]

    missing_1c = [
        r for r in results
        if r["Тип"] == "Нет в 1С"
    ]

    missing_esf = [
        r for r in results
        if r["Тип"] == "Нет в ЭСФ"
    ]

    return render_template(
        "results.html",
        results=results,
        total=total,
        ok_count=ok_count,
        error_count=error_count,
        match_percent=match_percent,
        matches=matches,
        differences=differences,
        duplicates=duplicates,
        missing_1c=missing_1c,
        missing_esf=missing_esf
    )


@app.route("/download")
def download():

    return send_file(
        "reports/report.xlsx",
        as_attachment=True
    )


if __name__ == "__main__":

    app.run(
        debug=True
    )