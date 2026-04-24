import time
import base64
import requests
import pandas as pd
from pathlib import Path

from config.settings import AIRTABLE_BASE_ID, AIRTABLE_TABLE_NAME


def normalize_text(s: str) -> str:
    return str(s).strip().lower()


def normalize_date(date_str: str) -> str:
    try:
        return pd.to_datetime(date_str).strftime("%Y-%m-%d")
    except Exception:
        return ""


def upload_to_airtable(file_path, course: str, pl: str, date_str: str, token: str, upload_logs: list, retries: int = 2):
    if not token:
        upload_logs.append({
            "Course": course, "PL": pl, "Date": date_str,
            "File": str(file_path), "Status": "Skipped", "Reason": "No token",
        })
        return

    url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}"
    headers = {"Authorization": f"Bearer {token}"}
    formula = (
        f"AND("
        f"LOWER({{Run Code - Identifier}}) = '{normalize_text(course)}',"
        f"LOWER({{PL Name}}) = '{normalize_text(pl)}',"
        f"DATETIME_FORMAT({{Date}}, 'YYYY-MM-DD') = '{normalize_date(date_str)}'"
        f")"
    )

    attempt = 0
    while attempt <= retries:
        try:
            res = requests.get(url, headers=headers, params={"filterByFormula": formula})
            records = res.json().get("records", [])

            if len(records) == 0:
                raise Exception("No matching record")
            if len(records) > 1:
                raise Exception("Multiple matching records")

            record = records[0]
            if record["fields"].get("Feedback Report"):
                upload_logs.append({
                    "Course": course, "PL": pl, "Date": date_str,
                    "File": str(file_path), "Status": "Skipped", "Reason": "Already uploaded",
                })
                return

            with open(file_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode()

            upload_url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}/{record['id']}"
            payload = {
                "fields": {
                    "Feedback Report": [{
                        "filename": Path(file_path).name,
                        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "data": encoded,
                    }]
                }
            }
            r = requests.patch(upload_url, json=payload, headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            })

            if r.status_code in (200, 201):
                upload_logs.append({
                    "Course": course, "PL": pl, "Date": date_str,
                    "File": str(file_path), "Status": "Uploaded", "Reason": "",
                })
                return
            else:
                raise Exception(r.text)

        except Exception as e:
            attempt += 1
            if attempt > retries:
                upload_logs.append({
                    "Course": course, "PL": pl, "Date": date_str,
                    "File": str(file_path), "Status": "Failed", "Reason": str(e),
                })
            else:
                time.sleep(1)
