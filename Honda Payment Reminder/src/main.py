import pandas as pd
import win32com.client as win32
import os
import traceback

# =========================
# CONFIG
# =========================
PREVIEW_MODE = False   # Always send directly

# =========================
# PATH HANDLING
# =========================
def get_file_path():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_dir = os.path.join(base_dir, "input")

    default_file = os.path.join(input_dir, "input.xlsx")
    if os.path.exists(default_file):
        return default_file

    if os.path.exists(input_dir):
        files = [f for f in os.listdir(input_dir) if f.endswith((".xlsx", ".xls"))]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(input_dir, x)), reverse=True)
            return os.path.join(input_dir, files[0])

    print("❌ No Excel file found in input folder")
    exit()

FILE_PATH = get_file_path()
print(f"📂 Using file: {FILE_PATH}")

# =========================
# LOAD DATA
# =========================
try:
    df = pd.read_excel(FILE_PATH)
except Exception as e:
    print("❌ Error reading Excel:", e)
    exit()

# =========================
# VALIDATE COLUMNS
# =========================
required_columns = [
    "Dealer Code", "Frame no", "Invoice no", "Invoice date",
    "Account Name", "Bank name", "Account number", "IFSC Code",
    "Refrence no.", "Payment date", "Claim Amount", "TO"
]

missing = [col for col in required_columns if col not in df.columns]
if missing:
    print(f"❌ Missing columns: {missing}")
    exit()

# =========================
# DATE FORMAT
# =========================
def format_date(value):
    dt = pd.to_datetime(value, dayfirst=True, errors='coerce')
    return dt.strftime('%d-%m-%Y') if not pd.isna(dt) else ''

# =========================
# IMPORT TEMPLATE
# =========================
from email_template import generate_email_html

# =========================
# OUTLOOK
# =========================
outlook = win32.Dispatch('outlook.application')

# =========================
# LOG FILE
# =========================
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
log_file = os.path.join(base_dir, "logs", "email_log.txt")
os.makedirs(os.path.dirname(log_file), exist_ok=True)

# =========================
# PROCESS
# =========================
sent_count = 0

for index, row in df.iterrows():
    try:
        if pd.isna(row['TO']):
            print(f"⚠️ Skipping row {index} (No TO)")
            continue

        invoice_date = format_date(row['Invoice date'])
        payment_date = format_date(row['Payment date'])

        html_body = generate_email_html(row, invoice_date, payment_date)

        mail = outlook.CreateItem(0)
        mail.To = str(row['TO'])
        mail.CC = str(row.get('CC', ''))
        mail.Subject = f"EW Plus Claim Payment - {row['Dealer Code']} - {row['Frame no']}"
        mail.HTMLBody = html_body

        # 🚀 DIRECT SEND (NO PREVIEW / NO DELAY)
        mail.Send()
        sent_count += 1

        print(f"✅ Sent: {row['Frame no']}")

    except Exception as e:
        error_msg = f"❌ Row {index} failed: {str(e)}"
        print(error_msg)

        with open(log_file, "a") as f:
            f.write(error_msg + "\n")
            f.write(traceback.format_exc() + "\n")

print(f"\n🚀 Done! Total sent: {sent_count}\n")