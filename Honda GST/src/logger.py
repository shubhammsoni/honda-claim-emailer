import os
import pandas as pd
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
LOG_FILE = os.path.join(OUTPUT_DIR, "logs.xlsx")


def ensure_output_dir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def log_entry(data):
    ensure_output_dir()

    row = dict(data)
    row["logged_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df = pd.DataFrame([row])

    if os.path.exists(LOG_FILE):
        existing = pd.read_excel(LOG_FILE)
        df = pd.concat([existing, df], ignore_index=True)

    df.to_excel(LOG_FILE, index=False)