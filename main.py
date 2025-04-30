import gradio as gr
import pandas as pd
import os
import zipfile
import requests
from concurrent.futures import ThreadPoolExecutor

# === CONFIGURATION ===
# Set your public Google Sheet ID here.
GOOGLE_SHEET_ID = "1bJQH3omGEju1mFR_AX5Fhk2KGWuC9ZvK5YuKf3AyOtA"
# Construct the CSV export URL for the published Google Sheet.
GOOGLE_SHEET_CSV_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=csv"

# Folder to store downloaded PDFs
PDF_FOLDER = "pdfs"
os.makedirs(PDF_FOLDER, exist_ok=True)

# === BACKEND LOGIC ===

def process(file):
    if file is None:
        return "No file uploaded.", None, None

    # --- Part 1: Extract APIR codes from uploaded Excel ---
    try:
        # Read the uploaded Excel; we assume it has a header row but we only care about column order
        user_df = pd.read_excel(file.name, header=0)
    except Exception as e:
        return f"Error reading the uploaded Excel: {e}", None, None

    # Extract the first column from the user's uploaded file as the list of APIR codes.
    # We convert them to string and strip any spaces.
    apir_codes = user_df.iloc[:, 0].astype(str).str.strip().tolist()

    # --- Part 2: Read the public Google Sheet (as CSV) ---
    try:
        gs_df = pd.read_csv(GOOGLE_SHEET_CSV_URL)
    except Exception as e:
        return f"Error reading the public Google Sheet: {e}", None, None

    # Assume the public Google Sheet has at least four columns with the following order:
    # Column 1: APIR code, Column 2: Product Name, Column 3: Date, Column 4: PDF URL
    # For output, we will only take the first four columns.

    # Prepare a list to store output rows as dictionaries.
    output_rows = []

    # For PDF downloads, store the download tasks information (for rows that are matched).
    download_tasks = []  # Each element will be a dict with keys: url and product_name

    # Loop over each APIR code from the user file.
    for code in apir_codes:
        # Search in the Google Sheet using the first column (regardless of header name)
        # We compare string versions of the APIR code.
        mask = gs_df.iloc[:, 0].astype(str).str.strip() == code
        if mask.any():
            # If a match is found, take the first match (should be unique)
            match_row = gs_df.loc[mask].iloc[0, :4]  # take only the first 4 columns
            row_dict = match_row.to_dict()
            # For PDF download: if the PDF URL is present in column 4, capture it.
            pdf_url = str(match_row.iloc[3]).strip() if not pd.isna(match_row.iloc[3]) else ""
            if pdf_url:
                # Product Name is the second column.
                product_name = str(match_row.iloc[1]).strip()
                download_tasks.append({"url": pdf_url, "product_name": product_name})
        else:
            # No match found: output row with the APIR code and blank columns for the rest.
            row_dict = {gs_df.columns[0]: code}
            # Ensure there are four columns in the output.
            for col_index in range(1, 4):
                # Use a blank string for missing columns.
                row_dict[f"Unnamed_{col_index}"] = ""
        output_rows.append(row_dict)

    # Create the output DataFrame.
    # For uniformity, set the columns to be (for example) "APIR Code", "Product Name", "Date", "PDF URL"
    output_df = pd.DataFrame(output_rows)
    # If the output DataFrame has more than 4 columns (because the Google Sheet might have more), take only the first 4.
    if output_df.shape[1] > 4:
        output_df = output_df.iloc[:, :4]
    # You can also rename the columns for clarity:
    output_df.columns = ["APIR Code", "Product Name", "Date", "PDF URL"]

    # Save the output DataFrame to an Excel file.
    output_excel = "output_results.xlsx"
    try:
        output_df.to_excel(output_excel, index=False)
    except Exception as e:
        return f"Error saving the output Excel: {e}", None, None

    # --- Part 3: Download PDFs and Zip them ---
    downloaded_files = []  # list of file paths for PDFs

    def download_pdf(task):
        url = task["url"]
        product_name = task["product_name"]
        # Create a safe filename: here we simply append ".pdf" after the product name.
        filename = f"{product_name} PDS.pdf"
        file_path = os.path.join(PDF_FOLDER, filename)
        try:
            response = requests.get(url)
            if response.status_code == 200:
                with open(file_path, "wb") as f:
                    f.write(response.content)
                downloaded_files.append(file_path)
        except Exception as e:
            # For now, we simply print the error and skip.
            print(f"Error downloading {url}: {e}")

    # Download PDFs in parallel using ThreadPoolExecutor.
    with ThreadPoolExecutor(max_workers=5) as executor:
        executor.map(download_pdf, download_tasks)

    # Zip the downloaded PDFs into one archive.
    zip_path = "pdfs_bundle.zip"
    try:
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for pdf_file in downloaded_files:
                zipf.write(pdf_file, arcname=os.path.basename(pdf_file))
    except Exception as e:
        return f"Error creating the ZIP archive: {e}", output_excel, None

    # --- End: Return outputs ---
    return "Processing complete!", output_excel, zip_path


# === FRONTEND UI ===
with gr.Blocks() as demo:
    gr.Markdown("## PDS Finder")
    gr.Markdown("Upload an Excel file. The file's first column (APIR codes) will be used to search a public Google Sheet.")
    
    with gr.Row():
        with gr.Column(scale=1, min_width=200):
            instructions = gr.Dropdown(label="Instructions", choices=[""], value="", interactive=False)
        with gr.Column(scale=5):
            file_upload = gr.File(label="Upload your Excel file", file_types=[".xlsx"])
    
    with gr.Row():
        clear_btn = gr.Button("Clear")
        submit_btn = gr.Button("Submit")

    with gr.Row():
        status_text = gr.Textbox(label="Processing Status", interactive=False)
    
    with gr.Row():
        excel_output = gr.File(label="Download Output Excel")
        zip_output = gr.File(label="Download ZIP of PDFs")

    # Hook up the Submit and Clear buttons.
    submit_btn.click(fn=process, inputs=[file_upload], outputs=[status_text, excel_output, zip_output])
    clear_btn.click(fn=lambda: ("", None, None), inputs=None, outputs=[status_text, excel_output, zip_output])

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 7860))  # Render gives us this PORT
    demo.launch(server_name="0.0.0.0", server_port=port)

