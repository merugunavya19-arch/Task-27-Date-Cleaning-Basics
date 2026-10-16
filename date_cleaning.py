"""Date Cleaning Basics — Veda Technology Data Analytics Internship.

Run from this project directory:
    python date_cleaning.py
"""
from pathlib import Path
from datetime import datetime
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "retail_sales_dataset.csv"
OUTPUT_FILE = BASE_DIR / "cleaned_retail_sales.csv"
ISSUE_FILE = BASE_DIR / "date_issue_log.csv"
DATE_FORMATS = [
    "%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y", "%B %d, %Y", "%Y/%m/%d",
    "%m.%d.%Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y", "%d-%m-%Y"
]


def parse_date(value):
    value = "" if value is None else str(value).strip()
    if not value:
        return None, "Missing date"
    for date_format in DATE_FORMATS:
        try:
            return datetime.strptime(value, date_format).date().isoformat(), ""
        except ValueError:
            continue
    return None, "Invalid date or unsupported format"


def clean_retail_dates(input_file=INPUT_FILE, output_file=OUTPUT_FILE, issue_file=ISSUE_FILE):
    data = pd.read_csv(input_file, dtype=str).fillna("")
    cleaned_rows, issues = [], []
    for _, row in data.iterrows():
        clean_date, issue = parse_date(row["sale_date"])
        record = row.to_dict()
        record["clean_sale_date"] = clean_date or ""
        record["date_status"] = "Valid" if not issue else "Issue"
        cleaned_rows.append(record)
        if issue:
            issues.append({
                "sale_id": row["sale_id"],
                "original_date": row["sale_date"],
                "issue_type": issue,
                "recommended_action": "Review source value and correct manually",
            })
    cleaned = pd.DataFrame(cleaned_rows)
    issue_log = pd.DataFrame(issues)
    cleaned.to_csv(output_file, index=False)
    issue_log.to_csv(issue_file, index=False)
    return cleaned, issue_log


if __name__ == "__main__":
    cleaned, issue_log = clean_retail_dates()
    print(f"Processed {len(cleaned)} rows.")
    print(f"Valid dates: {(cleaned['date_status'] == 'Valid').sum()}")
    print(f"Issues found: {len(issue_log)}")
    print(f"Saved: {OUTPUT_FILE}")
    print(f"Saved: {ISSUE_FILE}")
