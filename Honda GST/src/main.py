import os
import pandas as pd
from datetime import datetime

from formatter import format_table
from mailer import send_email
from logger import log_entry
from excel_generator import generate_excel

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_PATH = os.path.join(BASE_DIR, "data", "input.xlsx")
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "email_template.html")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

GST_SHEET = "GST"
MAIL_SHEET = "Mail"


def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


def clean_amount(col):
    return (
        col.fillna(0)
        .replace("-", 0)
        .astype(str)
        .str.replace(",", "")
        .replace("", "0")
        .astype(float)
    )


def format_amount(col):
    return col.map(lambda x: f"{x:,.2f}")


def build_email_map(email_df):
    email_df.columns = email_df.columns.str.strip()

    email_map = {}

    for _, row in email_df.iterrows():
        dealer = clean_text(row["Dealer code"])
        to_email = clean_text(row["To"])
        cc_raw = clean_text(row["CC"]).replace(" ", "")

        cc_list = cc_raw.split(",") if cc_raw else []

        email_map[dealer] = {
            "to": to_email,
            "cc": cc_list
        }

    return email_map


# ✅ FIXED SUMMARY FUNCTION
def write_summary(summary_rows, missing_rows):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    file = os.path.join(
        OUTPUT_DIR,
        f"run_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    with pd.ExcelWriter(file, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        pd.DataFrame(missing_rows).to_excel(
            writer,
            sheet_name="Missing",
            index=False
        )

    return file


def main():

    xls = pd.ExcelFile(FILE_PATH)
    print("Available sheets:", xls.sheet_names)

    df = pd.read_excel(FILE_PATH, sheet_name=GST_SHEET)
    email_df = pd.read_excel(FILE_PATH, sheet_name=MAIL_SHEET)

    df.columns = df.columns.str.strip()
    df["Dealer code"] = df["Dealer code"].astype(str).str.strip()

    df["Invoice Date"] = pd.to_datetime(df["Invoice Date"]).dt.strftime("%d-%m-%Y")

    for col in ["Taxable", "IGST", "CGST", "SGST"]:
        df[col] = clean_amount(df[col])

    email_map = build_email_map(email_df)

    grouped = df.groupby("Dealer code")

    summary = []
    missing = []

    with open(TEMPLATE_PATH) as f:
        template = f.read()

    for dealer, group in grouped:

        try:
            dealer = clean_text(dealer)

            # ❌ STRICT mapping
            if dealer not in email_map:
                missing.append({"Dealer": dealer})
                raise Exception("Dealer not found in Mail sheet")

            to_email = email_map[dealer]["to"]
            cc_emails = email_map[dealer]["cc"]

            if not to_email:
                raise Exception("Missing To Email")

            # Format for email
            email_group = group.copy()
            for col in ["Taxable", "IGST", "CGST", "SGST"]:
                email_group[col] = format_amount(email_group[col])

            # Excel attachment
            attachment = generate_excel(group.copy(), dealer)

            # HTML
            table_html = format_table(email_group)

            html_body = template.replace("{{table}}", table_html)\
                                .replace("{{name}}", group["Particulars"].iloc[0])

            # Send via Outlook
            success, error = send_email(
                to_email,
                cc_emails,
                "GST Input Credit Not Reflecting - Action Required",
                html_body,
                attachment
            )

            if not success:
                raise Exception(error)

            summary.append({
                "Dealer": dealer,
                "Email": to_email,
                "Status": "Sent"
            })

            log_entry({
                "dealer": dealer,
                "status": "sent"
            })

            print(f"✅ Sent: {dealer}")

        except Exception as e:
            summary.append({
                "Dealer": dealer,
                "Status": "Failed",
                "Error": str(e)
            })

            print(f"❌ Failed: {dealer} - {e}")

    file = write_summary(summary, missing)
    print("\nSummary saved at:", file)


if __name__ == "__main__":
    main()