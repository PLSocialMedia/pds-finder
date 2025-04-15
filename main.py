import gradio as gr

def process(file, apir_input):
    codes = apir_input.strip().split()
    return f"Processing {len(codes)} codes...", None, None

with gr.Blocks() as demo:
    gr.Markdown("## PDS Scraper & Validator")
    gr.Markdown("Upload an Excel file, run the scraper, download the results.")

    with gr.Row():
        with gr.Column(scale=1, min_width=200):
            instructions = gr.Dropdown(label="Instructions", choices=[""], value="", interactive=False)

        with gr.Column(scale=5):
            with gr.Row():
                file_upload = gr.File(label="Upload your Excel file here", file_types=[".xlsx"])
                apir_input = gr.Textbox(label="Enter APIR Codes (space separated)", placeholder="e.g. ABC123 DEF456 GHI789")

            with gr.Row():
                clear_btn = gr.Button("Clear")
                submit_btn = gr.Button("Submit", elem_id="submit-btn")

    with gr.Row():
        with gr.Column(scale=1):
            status_text = gr.Textbox(label="Processing Status", interactive=False)

            excel_output = gr.File(label="Processed Excel File")
            zip_output = gr.File(label="ZIP File (if available)")

    # Function hookups
    submit_btn.click(fn=process, inputs=[file_upload, apir_input], outputs=[status_text, excel_output, zip_output])
    clear_btn.click(fn=lambda: ("", None, None), inputs=None, outputs=[status_text, excel_output, zip_output])

if __name__ == "__main__":
    demo.launch()
