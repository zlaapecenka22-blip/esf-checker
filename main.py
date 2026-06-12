from checker import compare_files
from report_generator import create_report

results = compare_files(
    "data/1c.xlsx",
    "data/esf.xlsx"
)

create_report(results)

print(
    "Отчёт успешно создан: reports/report.xlsx"
)