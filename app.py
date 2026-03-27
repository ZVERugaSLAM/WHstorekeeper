import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai
import json
import tempfile
import re
import datetime
import io
import os
import time
import mimetypes

# Налаштування Gemini API через Streamlit Secrets
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=API_KEY)
except FileNotFoundError:
    st.error("Файл secrets.toml не знайдено.")
    st.stop()
except KeyError:
    st.error("Змінна GEMINI_API_KEY не знайдена у secrets.toml.")
    st.stop()

st.set_page_config(page_title="WHO Warehouse OCR & Automation", layout="wide")

def api_call_with_retry(func, *args, **kwargs):
    """Виконує API-виклик з очікуванням при 429 помилці."""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "Quota exceeded" in error_str:
                if attempt < max_retries - 1:
                    match = re.search(r'retry_delay\s*\{\s*seconds:\s*(\d+)\s*\}', error_str)
                    wait_time = int(match.group(1)) + 5 if match else 20
                    st.toast(f"⏳ Ліміт API. Очікування {wait_time} сек... (Спроба {attempt + 1}/{max_retries})")
                    time.sleep(wait_time)
                else:
                    raise e
            else:
                raise e

def process_document_with_gemini(file_path):
    try:
        uploaded_doc = api_call_with_retry(genai.upload_file, path=file_path)
        
        model = genai.GenerativeModel(model_name="gemini-2.5-flash")
        prompt = """
        You are a logistics assistant at WHO. Analyze the scanned warehouse documents.
        Extract the following data and return it STRICTLY in VALID JSON format.
        If information is missing, leave the value empty ("").

        Rules for specific fields:
        - "po_number": Extract ONLY the numeric part of the Purchase Order. Remove any "PO" prefixes. Example: "203793992".
        - "item_name_eng": Extract the FULL and complete English description of the item from the invoice or packing list.
        - "item_name_ukr": Provide an accurate Ukrainian translation of the "item_name_eng".
        - "exp_date": Format strictly as DD.MM.YYYY.
        - "supplier_name": Extract the supplier or vendor name from the document.
        - "invoice_info": Extract the Invoice number and date (e.g., "Invoice CI2600385.1 09.02.2026") or Waybill (Видаткова накладна) number and date.
        - "number_of_parcels": Extract the number of parcels/pallets/boxes (Кількість місць) from the document.

        JSON Keys to return:
        "act_number", "po_number", "item_name_eng", "item_name_ukr", "batch",
        "exp_date", "quantity", "number_of_parcels", "supplier_name", "invoice_info"
        """
        
        response = api_call_with_retry(model.generate_content, [prompt, uploaded_doc])
        
        json_text = response.text
        json_match = re.search(r'```json\n(.*?)\n```', json_text, re.DOTALL)
        if json_match:
            json_text = json_match.group(1)
            
        data = json.loads(json_text)
        
        manual_fields = ["or_number", "who_item_code", "project", "task", "award", "award_end_date", "donor", "requester", "wh"]
        for field in manual_fields:
            data[field] = ""
            
        return data
    except Exception as e:
        st.error(f"OCR Error: {e}")
        return {}

def process_packing_lists(file_paths):
    try:
        model = genai.GenerativeModel(model_name="gemini-2.5-flash")
        prompt = """
        You are a logistics data extraction assistant. Analyze packing list images representing multiple boxes on a pallet.
        Extract the general module information and ALL listed items across all provided pages.
        Always translate the English item description to Ukrainian for the 'item_description_ukr' field.
        If any data is missing on the document, leave the value empty ("").
        Dates must be strictly in DD.MM.YYYY format.

        Format the output STRICTLY as a valid JSON object with the following structure:
        {
          "module_name": "Extract the overall module or kit name (e.g., IEHK 2017, NCDK 2022 MODULE MEDICINES)",
          "module_batch": "Extract the overall module batch number (e.g., MPL00010680)",
          "items": [
            {
              "carton_no": "Carton or box number",
              "item_no": "Item code or number",
              "quantity": "Numeric quantity",
              "packing_unit": "UOM or packing unit (e.g., pack, vial, kit)",
              "item_description_ukr": "Accurate Ukrainian translation of the item description",
              "item_description_eng": "Full English item description",
              "batch_no": "Item batch number",
              "man_date": "Manufacturing date",
              "exp_date": "Expiry date"
            }
          ]
        }
        """
        
        content_request = [prompt]
        
        for fp in file_paths:
            mime_type, _ = mimetypes.guess_type(fp)
            if mime_type == 'application/pdf':
                uploaded_doc = api_call_with_retry(genai.upload_file, path=fp)
                content_request.append(uploaded_doc)
                time.sleep(2)
            else:
                with open(fp, "rb") as f:
                    doc_data = f.read()
                content_request.append({"mime_type": mime_type or "image/jpeg", "data": doc_data})
        
        response = api_call_with_retry(model.generate_content, content_request)
        
        json_text = response.text
        json_match = re.search(r'```json\n(.*?)\n```', json_text, re.DOTALL)
        if json_match:
            json_text = json_match.group(1)
            
        data = json.loads(json_text)
        return data
    except Exception as e:
        st.error(f"Packing List OCR Error: {e}")
        return {}

def set_cell_text(cell, text, bold=False):
    cell.text = ""
    run = cell.paragraphs[0].add_run(text)
    run.bold = bold

def generate_files_in_memory(data):
    base_name = f"PO_{data.get('po_number', 'XXX')}_Act_{data.get('act_number', 'XXX')}".replace("/", "-")
    
    excel_buffer = io.BytesIO()
    df = pd.DataFrame([{
        "ACT": data.get('act_number'), "PO": data.get('po_number'), "OR": data.get('or_number'),
        "Item Name [UKR]": data.get('item_name_ukr'), "Item Name [ENG]": data.get('item_name_eng'),
        "WHO Item code": data.get('who_item_code'), "Batch": data.get('batch'), 
        "Exp.date": data.get('exp_date'), "Quantity": data.get('quantity'), 
        "Project": data.get('project'), "Task": data.get('task'), "Award": data.get('award'), 
        "Award end date": data.get('award_end_date'), "Donor": data.get('donor'), 
        "Requester": data.get('requester'), "WH": data.get('wh')
    }])
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    
    word_buffer = io.BytesIO()
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(12)
    style.paragraph_format.line_spacing = 1.15
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run('WORLD HEALTH ORGANIZATION\n').bold = True
    p_title.runs[0].font.size = Pt(16)
    p_title.add_run('RECEIVING REPORT AND ACCEPTANCE INFORMATION').bold = True
    p_title.runs[1].font.size = Pt(14)
    
    doc.add_paragraph()
    
    header_table = doc.add_table(rows=6, cols=2)
    set_cell_text(header_table.cell(0, 0), "Purchase Order Number:")
    set_cell_text(header_table.cell(0, 1), f"{data.get('po_number', '')}")
    set_cell_text(header_table.cell(1, 0), "Registration Number:")
    set_cell_text(header_table.cell(1, 1), f"{data.get('invoice_info', '')}")
    set_cell_text(header_table.cell(2, 0), "Number of items Received:")
    set_cell_text(header_table.cell(2, 1), f"{data.get('quantity', '')}")
    set_cell_text(header_table.cell(3, 0), "Number of Parcels Received:")
    set_cell_text(header_table.cell(3, 1), f"{data.get('number_of_parcels', '')}")
    set_cell_text(header_table.cell(4, 0), "Item received (check one):")
    set_cell_text(header_table.cell(4, 1), "All [ ] ; Partial [ ]")
    set_cell_text(header_table.cell(5, 0), "Supplier name:")
    set_cell_text(header_table.cell(5, 1), f"{data.get('supplier_name', '')}")
    
    doc.add_paragraph()
    
    p_decl = doc.add_paragraph()
    p_decl.add_run('DECLARATION').bold = True
    doc.add_paragraph("The consignment against the above-mentioned PO has been received, invoice, packing list and other shipping documents are attached.")
    
    current_date = datetime.datetime.now().strftime("%d.%m.%Y")
    sign_table = doc.add_table(rows=4, cols=2)
    set_cell_text(sign_table.cell(0, 0), "Name of receiver:")
    set_cell_text(sign_table.cell(0, 1), "Oleksandr Tsybulnyk")
    set_cell_text(sign_table.cell(1, 0), "Title:")
    set_cell_text(sign_table.cell(1, 1), "Storekeeper")
    set_cell_text(sign_table.cell(2, 0), "Signature:")
    set_cell_text(sign_table.cell(2, 1), "")
    set_cell_text(sign_table.cell(3, 0), "Date:")
    set_cell_text(sign_table.cell(3, 1), current_date)
    
    doc.add_paragraph()
    
    p_remarks = doc.add_paragraph()
    p_remarks.add_run('Remarks:').bold = True
    
    if data.get('remark_desc'):
        table = doc.add_table(rows=2, cols=4)
        table.style = 'Table Grid'
        
        set_cell_text(table.rows[0].cells[0], 'Name, dosage form', bold=True)
        set_cell_text(table.rows[0].cells[1], 'Batch#', bold=True)
        set_cell_text(table.rows[0].cells[2], 'Quantity', bold=True)
        set_cell_text(table.rows[0].cells[3], 'Inconsistency description', bold=True)
        
        table.rows[1].cells[0].text = data.get('remark_item', '')
        table.rows[1].cells[1].text = data.get('remark_batch', '')
        table.rows[1].cells[2].text = data.get('remark_qty', '')
        table.rows[1].cells[3].text = data.get('remark_desc', '')
    else:
        doc.add_paragraph("None")

    doc.save(word_buffer)
    word_buffer.seek(0)
    
    return excel_buffer, word_buffer, base_name

# --- UI Додатку ---
st.title("📦 WHO Warehouse OCR & Automation")

tab1, tab2 = st.tabs(["📄 WRR Generator", "📦 Packing List OCR"])

with tab1:
    uploaded_file = st.file_uploader("Upload WRR document scan (PDF, JPG, PNG)", type=['pdf', 'png', 'jpg'], key="wrr_uploader")

    temp_file_path = None
    if uploaded_file is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as temp_file:
            temp_file.write(uploaded_file.read())
            temp_file_path = temp_file.name

        if st.button("Process Document (AI)"):
            with st.spinner("Analyzing document via Gemini..."):
                st.session_state['extracted_data'] = process_document_with_gemini(temp_file_path)
                
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)

    if 'extracted_data' in st.session_state and st.session_state['extracted_data']:
        st.subheader("Review and Edit Data")
        data = st.session_state['extracted_data']
        
        with st.form("data_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                act_number = st.text_input("ACT", value=data.get("act_number", ""))
                po_number = st.text_input("PO", value=data.get("po_number", ""))
                or_number = st.text_input("OR", value=data.get("or_number", ""))
                item_name_ukr = st.text_input("Item Name [UKR]", value=data.get("item_name_ukr", ""))
                item_name_eng = st.text_input("Item Name [ENG]", value=data.get("item_name_eng", ""))
                who_item_code = st.text_input("WHO Item code", value=data.get("who_item_code", ""))
            with col2:
                batch = st.text_input("Batch", value=data.get("batch", ""))
                exp_date = st.text_input("Exp.date (DD.MM.YYYY)", value=data.get("exp_date", ""))
                quantity = st.text_input("Quantity", value=data.get("quantity", ""))
                project = st.text_input("Project", value=data.get("project", ""))
                task = st.text_input("Task", value=data.get("task", ""))
            with col3:
                award = st.text_input("Award", value=data.get("award", ""))
                award_end_date = st.text_input("Award end date", value=data.get("award_end_date", ""))
                donor = st.text_input("Donor", value=data.get("donor", ""))
                requester = st.text_input("Requester", value=data.get("requester", ""))
                wh = st.text_input("WH (Warehouse)", value=data.get("wh", ""))

            st.markdown("---")
            st.subheader("WRR Specific Fields")
            col_w1, col_w2, col_w3 = st.columns(3)
            with col_w1:
                supplier_name = st.text_input("Supplier Name", value=data.get("supplier_name", ""))
            with col_w2:
                invoice_info = st.text_input("Invoice / Waybill", value=data.get("invoice_info", ""))
            with col_w3:
                number_of_parcels = st.text_input("Number of Parcels", value=data.get("number_of_parcels", ""))
            
            st.markdown("**Remarks / Discrepancies**")
            col_r1, col_r2, col_r3, col_r4 = st.columns(4)
            with col_r1: remark_item = st.text_input("Remark: Item Name", value="")
            with col_r2: remark_batch = st.text_input("Remark: Batch", value="")
            with col_r3: remark_qty = st.text_input("Remark: Quantity", value="")
            with col_r4: remark_desc = st.text_input("Remark: Inconsistency description", value="")

            submitted = st.form_submit_button("Generate Files")

        if submitted:
            final_data = {
                "act_number": act_number, "po_number": po_number, "or_number": or_number,
                "item_name_ukr": item_name_ukr, "item_name_eng": item_name_eng,
                "who_item_code": who_item_code, "batch": batch, "exp_date": exp_date,
                "quantity": quantity, "project": project, "task": task, "award": award,
                "award_end_date": award_end_date, "donor": donor, "requester": requester, 
                "wh": wh, "supplier_name": supplier_name, "invoice_info": invoice_info,
                "number_of_parcels": number_of_parcels,
                "remark_item": remark_item, "remark_batch": remark_batch, 
                "remark_qty": remark_qty, "remark_desc": remark_desc
            }
            
            excel_buffer, word_buffer, base_name = generate_files_in_memory(final_data)
            
            st.markdown("---")
            st.subheader("📥 Download Generated Files")
            col_btn1, col_btn2, col_empty = st.columns([1.5, 1.5, 7]) 
            with col_btn1:
                st.download_button(label="Download Excel Database", data=excel_buffer, file_name=f"{base_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col_btn2:
                st.download_button(label="Download WRR Document", data=word_buffer, file_name=f"WRR_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            
            st.markdown("---")
            st.subheader("Email Template (Copy and paste into Outlook)")
            st.text_input("Subject:", value=f"Warehouse Receiving Report / PO {po_number}")
            email_html = f"""
            <div style="font-family: Calibri, Arial, sans-serif; font-size: 14px;">
            <p>Dear Team,</p>
            <p>Please find attached the Warehouse Receiving Report (WRR) and Excel database for PO {po_number}.<br>
            Goods have been successfully received at the {wh} warehouse.<br>
            The data has been extracted and transferred to Farmasoft for balance registration.</p>
            <table border="1" style="border-collapse: collapse; width: 100%; max-width: 1000px; border-color: gray;">
            <tr style="background-color: rgba(128, 128, 128, 0.2);">
            <th style="padding: 8px; text-align: left;">PO</th><th style="padding: 8px; text-align: left;">Order request</th><th style="padding: 8px; text-align: left;">Requester</th><th style="padding: 8px; text-align: left;">Item Name [ENG]</th><th style="padding: 8px; text-align: left;">Batch</th><th style="padding: 8px; text-align: left;">Expiry date</th><th style="padding: 8px; text-align: left;">Total Quantity Inbound</th>
            </tr><tr>
            <td style="padding: 8px;">{po_number}</td><td style="padding: 8px;">{or_number}</td><td style="padding: 8px;">{requester}</td><td style="padding: 8px;">{item_name_eng}</td><td style="padding: 8px;">{batch}</td><td style="padding: 8px;">{exp_date}</td><td style="padding: 8px;">{quantity}</td>
            </tr></table></div>
            """
            st.markdown(email_html, unsafe_allow_html=True)

with tab2:
    uploaded_pl_files = st.file_uploader("Upload Packing List scans/photos (Multiple allowed)", type=['pdf', 'png', 'jpg'], accept_multiple_files=True, key="pl_uploader")
    
    if uploaded_pl_files and st.button("Extract to Excel", key="extract_pl"):
        with st.spinner("Processing packing lists via Gemini..."):
            temp_paths = []
            for f in uploaded_pl_files:
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{f.name.split('.')[-1]}") as temp_file:
                    temp_file.write(f.read())
                    temp_paths.append(temp_file.name)
            
            pl_data = process_packing_lists(temp_paths)
            
            for path in temp_paths:
                if os.path.exists(path):
                    os.remove(path)
            
            if pl_data and "items" in pl_data and len(pl_data["items"]) > 0:
                df_pl = pd.DataFrame(pl_data["items"])
                
                df_pl = df_pl.rename(columns={
                    "carton_no": "Carton no.",
                    "item_no": "Item no.",
                    "quantity": "Quantity",
                    "packing_unit": "Packing unit",
                    "item_description_ukr": "Item description UKR",
                    "item_description_eng": "Item description ENG",
                    "batch_no": "Batch no.",
                    "man_date": "Man. date",
                    "exp_date": "Exp. date"
                })
                
                cols = ["Carton no.", "Item no.", "Quantity", "Packing unit", "Item description UKR", "Item description ENG", "Batch no.", "Man. date", "Exp. date"]
                for col in cols:
                    if col not in df_pl.columns:
                        df_pl[col] = ""
                df_pl = df_pl[cols]

                st.success("Extraction complete!")
                st.dataframe(df_pl, use_container_width=True)
                
                excel_buffer_pl = io.BytesIO()
                with pd.ExcelWriter(excel_buffer_pl, engine='openpyxl') as writer:
                    df_pl.to_excel(writer, index=False, startrow=2)
                    worksheet = writer.sheets['Sheet1']
                    
                    module_name = pl_data.get('module_name', '')
                    module_batch = pl_data.get('module_batch', '')
                    batch_text = f"batch: {module_batch}" if module_batch else ""
                    
                    worksheet.cell(row=1, column=5, value=module_name)
                    worksheet.cell(row=2, column=5, value=batch_text)
                    
                excel_buffer_pl.seek(0)
                
                st.download_button(
                    label="Download Excel Document",
                    data=excel_buffer_pl,
                    file_name="Packing_List.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No items extracted or an error occurred.")