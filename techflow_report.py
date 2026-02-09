import os
import json
import csv
import smtplib
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from collections import Counter

from pyairtable import Api
from dotenv import load_dotenv
from  dashboard_html import generate_dashboard_html, save_dashboard

load_dotenv()

AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY")
AIRTABLE_BASE_ID = os.getenv("AIRTABLE_BASE_ID")
AIRTABLE_TABLE_NAME = "Requests"

SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_TO = os.getenv("EMAIL_TO")

REPORTS_DIR = "reports"


def fetch_all_requests():
    """
    Get all fields from requests table in Airtable.
    """
    api = Api(AIRTABLE_API_KEY)
    table = api.table(AIRTABLE_BASE_ID, AIRTABLE_TABLE_NAME)

    records = table.all()

    print(f"Отримано {len(records)} записів з Airtable")
    return records


def parse_datetime(value):
    """
    Converts a date string from Airtable to a datetime object.
    """
    if not value:
        return None

    formats = [
        "%Y-%m-%dT%H:%M:%S.%fZ",
        "%Y-%m-%dT%H:%M:%SZ",
        "%Y-%m-%dT%H:%M:%S.%f",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(value, fmt)
            return dt.replace(tzinfo=ZoneInfo("UTC"))
        except ValueError:
            continue

    return None


def calculate_metrics(records):
    """
    Calculates all metrics for the weekly report.
    Arguments:
        records (list[dict]): records from Airtable
    Returns:
        dict: structured report with metrics and details
    """
    now = datetime.now(ZoneInfo("Europe/Kyiv"))
    week_ago = now - timedelta(days=7)

    all_requests = []
    for record in records:
        fields = record.get("fields", {})

        created_at = parse_datetime(fields.get("Created At"))
        assigned_at = parse_datetime(fields.get("Assigned At"))
        closed_at = parse_datetime(fields.get("Closed At"))
        status = fields.get("Status", "Unknown")

        assignee_name = (fields.get("Assignee Name") or ["Unassigned"])[0]

        service = (fields.get("Service Name") or ["Unknown"])[0]

        all_requests.append({
            "request_id": fields.get("Request ID", "N/A"),
            "status": status,
            "assignee": assignee_name,
            "service": service,
            "description": fields.get("Description", ""),
            "created_at": created_at,
            "assigned_at": assigned_at,
            "closed_at": closed_at,
        })

    new_this_week = [
        r for r in all_requests
        if r["created_at"] and r["created_at"] >= week_ago
    ]
    new_requests_count = len(new_this_week)

    closed_this_week = [
        r for r in all_requests
        if r["closed_at"] and r["closed_at"] >= week_ago
    ]
    closed_requests_count = len(closed_this_week)

    processing_times = []
    for r in all_requests:
        if r["created_at"] and r["closed_at"]:
            delta = r["closed_at"] - r["created_at"]
            processing_times.append(delta.total_seconds() / 3600)

    avg_processing_time_hours = (
        round(sum(processing_times) / len(processing_times), 2)
        if processing_times else 0
    )

    closed_by_consultant = Counter(
        r["assignee"] for r in all_requests
        if r["status"] == "Closed" and r["assignee"] != "Unassigned"
    )
    top_3_consultants = closed_by_consultant.most_common(3)

    in_progress_count = sum(
        1 for r in all_requests if r["status"] == "In progress"
    )

    overdue_count = sum(
        1 for r in all_requests if r["status"] == "More then 24 hours"
    )

    reaction_times = []
    for r in all_requests:
        if r["created_at"] and r["assigned_at"]:
            delta = r["assigned_at"] - r["created_at"]
            reaction_times.append(delta.total_seconds() / 3600)

    avg_reaction_time_hours = (
        round(sum(reaction_times) / len(reaction_times), 2)
        if reaction_times else 0
    )

    service_distribution = Counter(r["service"] for r in all_requests)
    total_services = sum(service_distribution.values())
    service_stats = [
        {"service": s, "count": c, "percent": round(c * 100 / total_services, 1)}
        for s, c in service_distribution.most_common()
    ]

    open_by_consultant = Counter(
        r["assignee"] for r in all_requests
        if r["status"] not in ("Closed", "More then 24 hours") and r["assignee"] != "Unassigned"
    )

    total_requests = len(all_requests)

    report = {
        "report_date": now.strftime("%Y-%m-%d %H:%M") + " (Kyiv)",
        "period": f"{week_ago.strftime('%Y-%m-%d')} — {now.strftime('%Y-%m-%d')}",
        "metrics": {
            "new_requests_this_week": new_requests_count,
            "closed_requests_this_week": closed_requests_count,
            "avg_processing_time_hours": avg_processing_time_hours,
            "top_3_consultants": [
                {"name": name, "closed_count": count}
                for name, count in top_3_consultants
            ],
            "in_progress_count": in_progress_count,
            "overdue_count": overdue_count,
            "avg_reaction_time_hours": avg_reaction_time_hours,
            "total_requests": total_requests,
            "service_stats": service_stats,
            "consultant_workload": dict(open_by_consultant),
        },
        "details": {
            "new_requests": [
                {
                    "request_id": r["request_id"],
                    "service": r["service"],
                    "assignee": r["assignee"],
                    "status": r["status"],
                    "created_at": r["created_at"].strftime("%Y-%m-%d %H:%M") if r["created_at"] else None,
                }
                for r in new_this_week
            ],
            "closed_requests": [
                {
                    "request_id": r["request_id"],
                    "service": r["service"],
                    "assignee": r["assignee"],
                    "closed_at": r["closed_at"].strftime("%Y-%m-%d %H:%M") if r["closed_at"] else None,
                }
                for r in closed_this_week
            ],
        }
    }

    return report


def save_json(report, directory, filename):
    """
    Save report in JSON format.
    """
    filepath = os.path.join(directory, filename)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"JSON звіт збережено: {filepath}")
    return filepath


def save_csv(report, directory, filename):
    """
    Save report in CSV format.

    """
    filepath = os.path.join(directory, filename)
    metrics = report["metrics"]

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        writer.writerow(["TechFlow Consulting — Weekly Report"])
        writer.writerow(["Report Date", report["report_date"]])
        writer.writerow(["Period", report["period"]])
        writer.writerow([])

        writer.writerow(["--- ОСНОВНІ МЕТРИКИ ---"])
        writer.writerow(["Нових запитів за тиждень", metrics["new_requests_this_week"]])
        writer.writerow(["Закритих запитів за тиждень", metrics["closed_requests_this_week"]])
        writer.writerow(["Середній час обробки (годин)", metrics["avg_processing_time_hours"]])
        writer.writerow(["Середній час реакції (годин)", metrics["avg_reaction_time_hours"]])
        writer.writerow(["Запитів у роботі", metrics["in_progress_count"]])
        writer.writerow(["Прострочених запитів (>24г)", metrics["overdue_count"]])
        writer.writerow(["Загальна кількість запитів", metrics["total_requests"]])
        writer.writerow([])

        writer.writerow(["--- ТОП-3 КОНСУЛЬТАНТИ (закриті запити) ---"])
        writer.writerow(["Консультант", "Закритих запитів"])
        for consultant in metrics["top_3_consultants"]:
            writer.writerow([consultant["name"], consultant["closed_count"]])
        writer.writerow([])

        writer.writerow(["--- РОЗПОДІЛ ПО ПОСЛУГАХ ---"])
        writer.writerow(["Послуга", "Кількість запитів", "Відсоток від загальних запитів"])
        for service in metrics["service_stats"]:
            writer.writerow([
                service["service"],
                service["count"],
                f"{service['percent']}%",
            ])
        writer.writerow([])

        writer.writerow(["--- НАВАНТАЖЕННЯ КОНСУЛЬТАНТІВ (відкриті запити) ---"])
        writer.writerow(["Консультант", "Відкритих запитів"])
        for name, count in metrics["consultant_workload"].items():
            writer.writerow([name, count])
        writer.writerow([])

        writer.writerow(["--- НОВІ ЗАПИТИ ЗА ТИЖДЕНЬ ---"])
        writer.writerow(["Request ID", "Послуга", "Консультант", "Статус", "Створено"])
        for r in report["details"]["new_requests"]:
            writer.writerow([r["request_id"], r["service"], r["assignee"], r["status"], r["created_at"]])
        writer.writerow([])

        writer.writerow(["--- ЗАКРИТІ ЗАПИТИ ЗА ТИЖДЕНЬ ---"])
        writer.writerow(["Request ID", "Послуга", "Консультант", "Закрито"])
        for r in report["details"]["closed_requests"]:
            writer.writerow([r["request_id"], r["service"], r["assignee"], r["closed_at"]])

    print(f"CSV звіт збережено: {filepath}")
    return filepath


def format_email_body(report):
    """
    Generates HTML body of report email.

    Uses inline CSS for compatibility with email clients
    (Gmail, Outlook, etc. do not support external CSS files).
    Includes cards with key metrics and tables with details.
    """
    metrics = report["metrics"]

    top_consultants_rows = ""
    for i, c in enumerate(metrics["top_3_consultants"], 1):
        top_consultants_rows += f"<tr><td>{i}</td><td>{c['name']}</td><td>{c['closed_count']}</td></tr>"

    workload_rows = ""
    for name, count in metrics["consultant_workload"].items():
        workload_rows += f"<tr><td>{name}</td><td>{count}</td></tr>"

    service_rows = ""
    for service in metrics["service_stats"]:
        service_rows += f"<tr><td>{service['service']}</td><td>{service['count']}</td><td>{service['percent']}%</td></tr>"

    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; color: #333; }}
            h1 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }}
            h2 {{ color: #2c3e50; margin-top: 30px; }}
            table {{ border-collapse: collapse; width: 100%; margin: 10px 0; }}
            th, td {{ border: 1px solid #ddd; padding: 8px 12px; text-align: left; }}
            th {{ background-color: #3498db; color: white; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            .metric-card {{
                display: inline-block; background: #ecf0f1; padding: 15px 25px;
                margin: 5px; border-radius: 8px; text-align: center; min-width: 150px;
            }}
            .metric-value {{ font-size: 28px; font-weight: bold; color: #2c3e50; }}
            .metric-label {{ font-size: 12px; color: #7f8c8d; margin-top: 5px; }}
        </style>
    </head>
    <body>
        <h1>📊 TechFlow Consulting — Щотижневий звіт</h1>
        <p><strong>Період:</strong> {report['period']}<br>
        <strong>Дата звіту:</strong> {report['report_date']}</p>

        <h2>Основні метрики</h2>
        <div>
            <div class="metric-card">
                <div class="metric-value">{metrics['new_requests_this_week']}</div>
                <div class="metric-label">Нових запитів</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{metrics['closed_requests_this_week']}</div>
                <div class="metric-label">Закритих запитів</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{metrics['avg_processing_time_hours']}г</div>
                <div class="metric-label">Середній час обробки</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{metrics['avg_reaction_time_hours']}г</div>
                <div class="metric-label">Середній час реакції</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{metrics['overdue_count']}</div>
                <div class="metric-label">Прострочених (&gt;24г)</div>
            </div>
        </div>

        <h2>🏆 Топ-3 консультанти (закриті запити)</h2>
        <table>
            <tr><th>#</th><th>Консультант</th><th>Закритих запитів</th></tr>
            {top_consultants_rows if top_consultants_rows else '<tr><td colspan="3">Немає закритих запитів</td></tr>'}
        </table>

        <h2>📋 Навантаження консультантів (відкриті запити)</h2>
        <table>
            <tr><th>Консультант</th><th>Відкритих запитів</th></tr>
            {workload_rows if workload_rows else '<tr><td colspan="2">Немає відкритих запитів</td></tr>'}
        </table>

        <h2>Розподіл по послугах</h2>
        <table>
            <tr><th>Послуга</th><th>Кількість</th></tr>
            {service_rows}
        </table>

        <hr>
        <p style="color: #95a5a6; font-size: 12px;">
            Звіт згенеровано автоматично системою TechFlow Report Generator.<br>
            Детальні дані доступні у прикріплених файлах (CSV, JSON).
        </p>
    </body>
    </html>
    """
    return html


def send_email(report, csv_path, json_path, dashboard_path):
    """
    Sends a report to the email manager.

    The email contains:
    - HTML body with metrics and tables
    - Attached CSV file
    - Attached JSON file
    - Attached interactive HTML dashboard
    """
    msg = MIMEMultipart()
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg["Subject"] = f"📊 TechFlow — Щотижневий звіт ({report['period']})"

    html_body = format_email_body(report)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    for filepath in [csv_path, json_path, dashboard_path]:
        with open(filepath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(filepath)}")
            msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        print(f"Звіт надіслано на {EMAIL_TO}")
    except Exception as e:
        print(f"Помилка відправки email: {e}")
        print("Перевірте SMTP налаштування у .env файлі")


def main():
    """
    Entry point into the application. Performs the following in sequence:
    1. Check configuration (.env variables)
    2. Get data from Airtable
    3. Calculate metrics
    4. Output results to the console
    5. Save to JSON and CSV
    6. Send to email
    """
    print("=" * 60)
    print("  TechFlow Consulting — Weekly Report Generator")
    print("=" * 60)
    print()

    required_vars = {
        "AIRTABLE_API_KEY": AIRTABLE_API_KEY,
        "AIRTABLE_BASE_ID": AIRTABLE_BASE_ID,
    }
    missing = [name for name, value in required_vars.items() if not value]

    if missing:
        print(f"Відсутні обов'язкові змінні середовища: {', '.join(missing)}")
        print("Створіть .env файл за прикладом .env.example")
        print("Або встановіть змінні: export AIRTABLE_API_KEY=pat_xxx...")
        return

    email_vars = [SMTP_HOST, SMTP_USER, SMTP_PASSWORD, EMAIL_FROM, EMAIL_TO]
    email_configured = all(email_vars)
    if not email_configured:
        print("Email не налаштований — звіт буде збережений тільки у файли")
        print()

    print("Отримання даних з Airtable...")
    try:
        records = fetch_all_requests()
    except Exception as e:
        print(f"Помилка підключення до Airtable: {e}")
        print("Перевірте AIRTABLE_API_KEY та AIRTABLE_BASE_ID у .env")
        return

    if not records:
        print("Таблиця Requests порожня — немає даних для звіту")
        return

    print("Розрахунок метрик...")
    report = calculate_metrics(records)

    metrics = report["metrics"]
    print()
    print(f"  Період: {report['period']}")
    print(f"  Нових запитів за тиждень:    {metrics['new_requests_this_week']}")
    print(f"  Закритих запитів за тиждень:  {metrics['closed_requests_this_week']}")
    print(f"  Середній час обробки:         {metrics['avg_processing_time_hours']} годин")
    print(f"  Середній час реакції:          {metrics['avg_reaction_time_hours']} годин")
    print(f"  У роботі зараз:               {metrics['in_progress_count']}")
    print(f"  Прострочених (>24г):          {metrics['overdue_count']}")
    print(f"  Загальна кількість:            {metrics['total_requests']}")
    print()

    if metrics["top_3_consultants"]:
        print("  🏆 Топ-3 консультанти:")
        for i, c in enumerate(metrics["top_3_consultants"], 1):
            print(f"     {i}. {c['name']} — {c['closed_count']} закритих")
    else:
        print("  🏆 Топ-3 консультанти: немає закритих запитів за період")
    print()


    date_str = datetime.now(ZoneInfo("Europe/Kyiv")).strftime("%Y-%m-%d")
    date_dir = os.path.join(REPORTS_DIR, date_str)
    os.makedirs(date_dir, exist_ok=True)

    json_path = save_json(report, date_dir, f"techflow_report_{date_str}.json")
    csv_path = save_csv(report, date_dir, f"techflow_report_{date_str}.csv")

    dashboard_html = generate_dashboard_html(report)
    dashboard_path = save_dashboard(dashboard_html, date_dir, f"techflow_dashboard_{date_str}.html")

    if email_configured:
        print()
        print("Надсилання звіту на email...")
        send_email(report, csv_path, json_path, dashboard_path)
    else:
        print()
        print("Email не налаштований — пропускаємо відправку")

    print()
    print("Готово!")


if __name__ == "__main__":
    main()

