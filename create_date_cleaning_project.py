from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

ROOT = Path('/home/ubuntu/Date-Cleaning-Basics')
ROOT.mkdir(parents=True, exist_ok=True)

# Retail sales data with intentionally inconsistent and invalid date formats.
rows = [
    ['S001','2024-01-15','Aarav Sharma','Electronics',1299.00],
    ['S002','15/02/2024','Meera Patel','Furniture',450.50],
    ['S003','03-20-2024','Rohan Kumar','Office Supplies',89.99],
    ['S004','April 5, 2024','Ananya Singh','Electronics',799.00],
    ['S005','2024/05/18','Vikram Rao','Furniture',325.75],
    ['S006','06.22.2024','Diya Nair','Office Supplies',55.00],
    ['S007','July 1 2024','Karan Shah','Electronics',149.99],
    ['S008','2024-08-30','Sneha Iyer','Furniture',680.00],
    ['S009','31/09/2024','Rahul Reddy','Office Supplies',42.50],
    ['S010','2024/10/12','Priya Desai','Electronics',999.00],
    ['S011','November 3, 2024','Aditya Verma','Furniture',210.25],
    ['S012','2024-02-29','Kavya Menon','Office Supplies',75.00],
    ['S013','13/13/2024','Sameer Khan','Electronics',250.00],
    ['S014','','Neha Joshi','Furniture',390.00],
    ['S015','2024-12-21','Amit Gupta','Office Supplies',65.40],
    ['S016','Jan 7, 2025','Isha Rao','Electronics',525.00],
    ['S017','2025/02/14','Nikhil Das','Furniture',275.25],
    ['S018','30-03-2025','Pooja Shah','Office Supplies',33.99],
    ['S019','April 31, 2025','Vivek Nair','Electronics',875.00],
    ['S020','2025-05-26','Riya Mehta','Furniture',410.00],
],
columns = ['sale_id','sale_date','customer_name','category','sales_amount']
raw = pd.DataFrame(rows[0], columns=columns)
raw.to_csv(ROOT/'retail_sales_dataset.csv', index=False)

# Explicit formats prevent ambiguous date parsing and preserve an issue log.
DATE_FORMATS = [
    '%Y-%m-%d', '%d/%m/%Y', '%m-%d-%Y', '%B %d, %Y', '%Y/%m/%d',
    '%m.%d.%Y', '%B %d %Y', '%b %d, %Y', '%b %d %Y', '%d-%m-%Y'
]

def parse_date(value):
    value = '' if value is None else str(value).strip()
    if not value:
        return None, 'Missing date'
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date().isoformat(), ''
        except ValueError:
            continue
    return None, 'Invalid date or unsupported format'

cleaned_rows = []
issues = []
for _, row in raw.iterrows():
    clean_date, issue = parse_date(row['sale_date'])
    out = row.to_dict()
    out['clean_sale_date'] = clean_date or ''
    out['date_status'] = 'Valid' if not issue else 'Issue'
    cleaned_rows.append(out)
    if issue:
        issues.append({
            'sale_id': row['sale_id'],
            'original_date': row['sale_date'],
            'issue_type': issue,
            'recommended_action': 'Review source value and correct manually'
        })
cleaned = pd.DataFrame(cleaned_rows)
issue_df = pd.DataFrame(issues)
cleaned.to_csv(ROOT/'cleaned_retail_sales.csv', index=False)
issue_df.to_csv(ROOT/'date_issue_log.csv', index=False)

script = '''"""Date Cleaning Basics — Veda Technology Data Analytics Internship.

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
'''
(ROOT/'date_cleaning.py').write_text(script, encoding='utf-8')

# Excel issue log for convenient review.
wb = Workbook(); ws = wb.active; ws.title = 'Issue_Log'; sumws = wb.create_sheet('Summary')
navy='17365D'; blue='2F75B5'; light='D9EAF7'; pale='F4F8FB'; red='FCE4D6'; green='E2F0D9'; white='FFFFFF'; gray='666666'
def title(sheet,text,subtitle,end_col):
    sheet.merge_cells(start_row=1,start_column=1,end_row=1,end_column=end_col)
    c=sheet.cell(1,1,text); c.font=Font(size=16,bold=True,color=white); c.fill=PatternFill('solid',fgColor=navy)
    sheet.merge_cells(start_row=2,start_column=1,end_row=2,end_column=end_col)
    c=sheet.cell(2,1,subtitle); c.font=Font(size=10,italic=True,color=gray); c.alignment=Alignment(wrap_text=True)
def style_header(sheet,row,n):
    for col in range(1,n+1):
        c=sheet.cell(row,col); c.font=Font(bold=True,color=white); c.fill=PatternFill('solid',fgColor=blue); c.alignment=Alignment(horizontal='center',wrap_text=True)
def set_widths(sheet,n):
    for col in range(1,n+1):
        letter=get_column_letter(col); maxlen=max((len(str(c.value)) for c in sheet[letter] if c.value is not None),default=10)
        sheet.column_dimensions[letter].width=min(max(maxlen+2,14),48)

title(ws,'Date Issue Log','Invalid and missing dates identified during explicit parsing.',4)
for j,c in enumerate(issue_df.columns,1): ws.cell(4,j,c.replace('_',' ').title())
style_header(ws,4,4)
for i,row in enumerate(issue_df.itertuples(index=False,name=None),5):
    for j,v in enumerate(row,1): ws.cell(i,j,v)
ref=f'A4:D{4+len(issue_df)}'; tab=Table(displayName='DateIssueLog',ref=ref); tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True,showColumnStripes=False); ws.add_table(tab)
ws.freeze_panes='A5'; set_widths(ws,4)
for row in range(5,5+len(issue_df)):
    for col in range(1,5): ws.cell(row,col).fill=PatternFill('solid',fgColor=red)

title(sumws,'Date Cleaning Summary','Retail sales date standardization results.',3)
summary_rows=[('Input records',len(raw)),('Valid dates',int((cleaned['date_status']=='Valid').sum())),('Date issues',len(issue_df)),('Output format','ISO YYYY-MM-DD'),('Parser approach','Explicit list of accepted date formats'),('Files generated','retail_sales_dataset.csv, cleaned_retail_sales.csv, date_cleaning.py, date_issue_log.csv')]
for i,(k,v) in enumerate(summary_rows,4):
    sumws.cell(i,1,k).font=Font(bold=True,color=navy); sumws.cell(i,1).fill=PatternFill('solid',fgColor=light); sumws.cell(i,2,v).fill=PatternFill('solid',fgColor=pale); sumws.cell(i,2).alignment=Alignment(wrap_text=True)
sumws.column_dimensions['A'].width=28; sumws.column_dimensions['B'].width=92
for s in wb.worksheets: s.sheet_view.showGridLines=False
wb.save(ROOT/'date_issue_log.xlsx')
print(ROOT)
print('rows',len(raw),'valid',int((cleaned['date_status']=='Valid').sum()),'issues',len(issue_df))
