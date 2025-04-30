import gradio as gr
import pandas as pd
import os, zipfile, requests
from concurrent.futures import ThreadPoolExecutor

# === CONFIGURATION ===
GOOGLE_SHEET_ID      = "1bJQH3omGEju1mFR_AX5Fhk2KGWuC9ZvK5YuKf3AyOtA"
GOOGLE_SHEET_CSV_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=csv"

PDF_FOLDER = "pdfs"
os.makedirs(PDF_FOLDER, exist_ok=True)

# === BACK-END ===
def process(codes_text:str, file):          # <-- accept *both* inputs
    # ------------------------------------
    # 1. Collect the APIR codes requested
    # ------------------------------------
    apir_codes = []

    # -- Option A:  Excel upload -----
    if file is not None:
        try:
            user_df   = pd.read_excel(file.name, header=0)
            apir_codes = user_df.iloc[:, 0].astype(str).str.strip().tolist()
        except Exception as e:
            return f"Error reading the uploaded Excel: {e}", None, None

    # -- Option B:  codes pasted in box --
    elif codes_text.strip():
        # Split on newline / comma / space, then strip blanks
        raw = [p.strip() for p in re.split(r"[,\n\s]+", codes_text)]
        apir_codes = [c for c in raw if c]        # remove empties

    # Neither provided  → error
    else:
        return "Please upload an Excel file **or** paste APIR codes.", None, None

    # ------------------------------------
    # 2. Read Google Sheet & build results
    # ------------------------------------
    try:
        gs_df = pd.read_csv(GOOGLE_SHEET_CSV_URL)
    except Exception as e:
        return f"Error reading the public Google Sheet: {e}", None, None

    output_rows, download_tasks = [], []
    for code in apir_codes:
        mask = gs_df.iloc[:, 0].astype(str).str.strip() == code
        if mask.any():
            match = gs_df.loc[mask].iloc[0, :4]
            output_rows.append(match.to_dict())
            pdf_url = str(match.iloc[3]).strip() if not pd.isna(match.iloc[3]) else ""
            if pdf_url:
                download_tasks.append(
                    {"url": pdf_url, "product_name": str(match.iloc[1]).strip()}
                )
        else:
            output_rows.append(
                {gs_df.columns[0]: code,
                 "Unnamed_1": "", "Unnamed_2": "", "Unnamed_3": ""}
            )

    output_df = pd.DataFrame(output_rows).iloc[:, :4]
    output_df.columns = ["APIR Code", "Product Name", "Date", "PDF URL"]

    output_excel = "output_results.xlsx"
    output_df.to_excel(output_excel, index=False)

    # ------------------------------------
    # 3. Download PDFs  →  ZIP
    # ------------------------------------
    downloaded = []
    def dl(task):
        fn = os.path.join(PDF_FOLDER, f"{task['product_name']} PDS.pdf")
        try:
            r = requests.get(task["url"])
            if r.status_code == 200:
                with open(fn, "wb") as f: f.write(r.content)
                downloaded.append(fn)
        except Exception as e:
            print(f"Download failed {task['url']}: {e}")

    with ThreadPoolExecutor(max_workers=5) as ex:
        ex.map(dl, download_tasks)

    zip_path = "pdfs_bundle.zip"
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for pdf in downloaded:
            zipf.write(pdf, arcname=os.path.basename(pdf))

    return "Processing complete!", output_excel, zip_path


# === FRONT-END ===
with gr.Blocks() as demo:
    gr.Markdown("## PDS Finder")

    with gr.Tab("Upload Excel"):
        file_input = gr.File(label="Upload your Excel file", file_types=[".xlsx"])

    with gr.Tab("Paste APIR Codes"):
        codes_box = gr.Textbox(
            label="Paste APIR codes (comma, space, or newline separated)",
            placeholder="ABC123, DEF456 …",
            lines=4
        )

    with gr.Row():
        submit_btn = gr.Button("Run")
        clear_btn  = gr.Button("Clear")

    status  = gr.Textbox(label="Status", interactive=False)
    xlsx_o  = gr.File(label="Download Output Excel")
    zip_o   = gr.File(label="Download ZIP of PDFs")

    submit_btn.click(
        fn=process,
        inputs=[codes_box, file_input],      # order matches process()
        outputs=[status, xlsx_o, zip_o]
    )
    clear_btn.click(
        fn=lambda: ("", None, None, None),
        inputs=None,
        outputs=[status, xlsx_o, zip_o, file_input]
    )

if __name__ == "__main__":
    import re, os
    port = int(os.environ.get("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
