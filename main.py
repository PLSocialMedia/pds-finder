import os
import re
import json
import datetime
import zipfile
import requests
import pandas as pd
import gradio as gr
from concurrent.futures import ThreadPoolExecutor

# === GOOGLE SHEETS & FORMATTING SETUP ===
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import format_cell_range, CellFormat, Color

SCOPES    = ["https://www.googleapis.com/auth/spreadsheets"]
# Load service-account JSON from the environment variable:
sa_info   = json.loads(os.environ["GOOGLE_SA_JSON"])
creds     = ServiceAccountCredentials.from_json_keyfile_dict(sa_info, SCOPES)
gc        = gspread.authorize(creds)

GOOGLE_SHEET_ID = "1bJQH3omGEju1mFR_AX5Fhk2KGWuC9ZvK5YuKf3AyOtA"
sheet           = gc.open_by_key(GOOGLE_SHEET_ID).sheet1
HIGHLIGHT       = CellFormat(backgroundColor=Color(1, 1, 0.6))   # Pale yellow

# === ADMIN ALERT (via SendGrid) ===
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

ADMIN_EMAIL      = "sharukesh.seker@planlogic.com.au"            # <-- replace with your email
SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY")

def alert_admin(new_codes: list[str]):
    if not new_codes or not SENDGRID_API_KEY:
        return
    message = Mail(
        from_email=ADMIN_EMAIL,
        to_emails=ADMIN_EMAIL,
        subject=f"[PDS] {len(new_codes)} new APIR code(s) added",
        plain_text_content="New APIR codes:\n" + "\n".join(new_codes)
    )
    SendGridAPIClient(SENDGRID_API_KEY).send(message)

# === CONFIGURATION ===
GOOGLE_SHEET_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}"
    "/export?format=csv"
)
PDF_FOLDER = "pdfs"
os.makedirs(PDF_FOLDER, exist_ok=True)


# === BACK-END PROCESSING ===
def process(codes_text: str, file):
    # 1) Gather APIR codes from Excel upload or pasted text
    if file is not None:
        try:
            user_df   = pd.read_excel(file.name, header=0)
            apir_codes = user_df.iloc[:,0].astype(str).str.strip().tolist()
        except Exception as e:
            return f"Error reading Excel: {e}", None, None
    else:
        raw = re.split(r"[,\n\s]+", codes_text or "")
        apir_codes = [c.strip() for c in raw if c.strip()]
        if not apir_codes:
            return "Please upload Excel or paste APIR codes.", None, None

    # 2) Load your public Google Sheet data
    try:
        gs_df = pd.read_csv(GOOGLE_SHEET_CSV_URL)
    except Exception as e:
        return f"Error reading Google Sheet CSV: {e}", None, None

    # 3) Detect which codes are already in your “DB” sheet
    db_codes = set(sheet.col_values(1))
    new_codes_for_db = []

    output_rows   = []
    download_tasks = []

    for code in apir_codes:
        mask = gs_df.iloc[:,0].astype(str).str.strip() == code
        if mask.any():
            match = gs_df.loc[mask].iloc[0,:4]
            output_rows.append(match.to_dict())
            pdf_url = str(match.iloc[3]).strip() if not pd.isna(match.iloc[3]) else ""
            if pdf_url:
                download_tasks.append({
                    "url": pdf_url,
                    "product_name": str(match.iloc[1]).strip()
                })
        else:
            # Not found → blank row in output + track for DB insert
            output_rows.append({
                "APIR Code": code,
                "Product Name": "",
                "Date": "",
                "PDF URL": ""
            })
            if code not in db_codes:
                new_codes_for_db.append(code)

    # 4) Build output DataFrame and save to Excel
    output_df = pd.DataFrame(output_rows).iloc[:,:4]
    output_df.columns = ["APIR Code", "Product Name", "Date", "PDF URL"]
    output_excel = "output_results.xlsx"
    try:
        output_df.to_excel(output_excel, index=False)
    except Exception as e:
        return f"Error saving Excel: {e}", None, None

    # 5) Download matched PDFs in parallel
    downloaded = []
    def dl(task):
        fn = os.path.join(PDF_FOLDER, f"{task['product_name']} PDS.pdf")
        try:
            r = requests.get(task["url"])
            if r.status_code == 200:
                with open(fn, "wb") as f:
                    f.write(r.content)
                downloaded.append(fn)
        except Exception:
            pass

    with ThreadPoolExecutor(max_workers=5) as ex:
        ex.map(dl, download_tasks)

    zip_path = "pdfs_bundle.zip"
    try:
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for pdf in downloaded:
                zipf.write(pdf, arcname=os.path.basename(pdf))
    except Exception as e:
        return f"Error creating ZIP: {e}", output_excel, None

    # 6) Append new APIRs, then highlight exactly the bottom N rows
    if new_codes_for_db:
        today = datetime.date.today().isoformat()
        rows = [[code, "", "", "", f"Added {today}"]
                for code in new_codes_for_db]

        # 6a) append them (no explicit range, so gspread auto-appends)
        sheet.append_rows(rows, value_input_option="USER_ENTERED")

        # 6b) figure out the true last row number
        all_vals   = sheet.get_all_values()
        total_rows = len(all_vals)
        first_new  = total_rows - len(rows) + 1
        last_new   = total_rows

        # 6c) highlight those new rows in column A
        fmt_range = f"A{first_new}:A{last_new}"
        format_cell_range(sheet, fmt_range, HIGHLIGHT)

        # 6d) ping the admin
        alert_admin(new_codes_for_db)


    return "Processing complete!", output_excel, zip_path


# === GRADIO FRONT-END ===
with gr.Blocks() as demo:
    gr.Markdown("## PDS Finder — with Auto-DB Sync & Alerts")

    with gr.Tab("Upload Excel"):
        file_input = gr.File(label="Upload your Excel file",
                             file_types=[".xlsx"])
    with gr.Tab("Paste APIR Codes"):
        codes_box = gr.Textbox(
            label="Paste APIR codes",
            placeholder="ABC123, DEF456 …",
            lines=4
        )

    submit_btn = gr.Button("Run")
    clear_btn  = gr.Button("Clear")

    status    = gr.Textbox(label="Status", interactive=False)
    excel_out = gr.File(label="Download Output Excel")
    zip_out   = gr.File(label="Download ZIP of PDFs")

    submit_btn.click(
        fn=process,
        inputs=[codes_box, file_input],
        outputs=[status, excel_out, zip_out]
    )
    clear_btn.click(
        fn=lambda: ("", None, None),
        inputs=None,
        outputs=[status, excel_out, zip_out]
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
