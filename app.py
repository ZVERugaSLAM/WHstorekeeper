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
import os # Додано імпорт модуля os

# Налаштування Gemini API через Streamlit Secrets
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=API_KEY)
except FileNotFoundError:
    st.error("Файл secrets.toml не знайдено. Додай його для локальної роботи або налаштуй Secrets у Streamlit Cloud.")
    st.stop()
except KeyError:
    st.error("Змінна GEMINI_API_KEY не знайдена у secrets.toml.")
    st.stop()

st.set_page_config(page_title="WHO Warehouse OCR & Automation", layout="wide")

def process_document_with_gemini(file_path):
    try:
        uploaded_doc = genai.upload_file(path=file_path)
        
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
        "act_number",
        "po_number",
        "item_name_eng",
        "item_name_ukr",
        "batch",
        "exp_date",
        "quantity",
        "number_of_parcels",
        "supplier_name",
        "invoice_info"
        """
        
        response = model.generate_content([prompt, uploaded_doc])
        
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
        return {key: "" for key in ["act_number", "po_number", "item_name_ukr", "item_name_eng", "batch", "exp_date", "quantity", "number_of_parcels", "supplier_name", "invoice_info", "or_number", "who_item_code", "project", "task", "award", "award_end_date", "donor", "requester", "wh"]}

def set_cell_text(cell, text, bold=False):
    cell.text = ""
    run = cell.paragraphs[0].add_run(text)
    run.bold = bold

def generate_files_in_memory(data):
    """Генерує Excel та Word у буферах пам'яті (BytesIO) та повертає їх."""
    base_name = f"PO_{data['po_number']}_Act_{data['act_number']}".replace("/", "-")
    
    # Генерація Excel у пам'яті
    excel_buffer = io.BytesIO()
    df = pd.DataFrame([{
        "ACT": data['act_number'], "PO": data['po_number'], "OR": data['or_number'],
        "Item Name [UKR]": data['item_name_ukr'], "Item Name [ENG]": data['item_name_eng'],
        "WHO Item code": data['who_item_code'], "Batch": data['batch'], 
        "Exp.date": data['exp_date'], "Quantity": data['quantity'], 
        "Project": data['project'], "Task": data['task'], "Award": data['award'], 
        "Award end date": data['award_end_date'], "Donor": data['donor'], 
        "Requester": data['requester'], "WH": data['wh']
    }])
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    
    # Генерація Word у пам'яті
    word_buffer = io.BytesIO()
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(12)
    style.paragraph_format.line_spacing = 1.15
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title1 = p_title.add_run('WORLD HEALTH ORGANIZATION\n')
    r_title1.bold = True
    r_title1.font.size = Pt(16)
    r_title2 = p_title.add_run('RECEIVING REPORT AND ACCEPTANCE INFORMATION')
    r_title2.bold = True
    r_title2.font.size = Pt(14)
    
    doc.add_paragraph()
    
    header_table = doc.add_table(rows=6, cols=2)
    set_cell_text(header_table.cell(0, 0), "Purchase Order Number:")
    set_cell_text(header_table.cell(0, 1), f"{data['po_number']}")
    set_cell_text(header_table.cell(1, 0), "Registration Number:")
    set_cell_text(header_table.cell(1, 1), f"{data.get('invoice_info', '')}")
    set_cell_text(header_table.cell(2, 0), "Number of items Received:")
    set_cell_text(header_table.cell(2, 1), f"{data['quantity']}")
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
        
        hdr_cells = table.rows[0].cells
        set_cell_text(hdr_cells[0], 'Name, dosage form', bold=True)
        set_cell_text(hdr_cells[1], 'Batch#', bold=True)
        set_cell_text(hdr_cells[2], 'Quantity', bold=True)
        set_cell_text(hdr_cells[3], 'Inconsistency description', bold=True)
        
        row_cells = table.rows[1].cells
        row_cells[0].text = data.get('remark_item', '')
        row_cells[1].text = data.get('remark_batch', '')
        row_cells[2].text = data.get('remark_qty', '')
        row_cells[3].text = data.get('remark_desc', '')
    else:
        doc.add_paragraph("None")

    doc.save(word_buffer)
    word_buffer.seek(0)
    
    return excel_buffer, word_buffer, base_name

# --- UI Додатку ---
st.title("📦 WHO Warehouse OCR & Automation")

uploaded_file = st.file_uploader("Upload document scan (PDF, JPG, PNG)", type=['pdf', 'png', 'jpg'])

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

if 'extracted_data' in st.session_state:
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
        st.subheader("WRR Specific Fields (Not included in Excel)")
        col_w1, col_w2, col_w3 = st.columns(3)
        with col_w1:
            supplier_name = st.text_input("Supplier Name", value=data.get("supplier_name", ""))
        with col_w2:
            invoice_info = st.text_input("Invoice / Waybill (Registration Number)", value=data.get("invoice_info", ""))
        with col_w3:
            number_of_parcels = st.text_input("Number of Parcels", value=data.get("number_of_parcels", ""))
        
        st.markdown("**Remarks / Discrepancies (Fill only if needed)**")
        col_r1, col_r2, col_r3, col_r4 = st.columns(4)
        with col_r1:
            remark_item = st.text_input("Remark: Item Name", value="")
        with col_r2:
            remark_batch = st.text_input("Remark: Batch", value="")
        with col_r3:
            remark_qty = st.text_input("Remark: Quantity", value="")
        with col_r4:
            remark_desc = st.text_input("Remark: Inconsistency description", value="")

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
        
        # Відображення кнопок завантаження
        st.markdown("---")
        st.subheader("📥 Download Generated Files")
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            st.download_button(
                label="Download Excel Database",
                data=excel_buffer,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_btn2:
            st.download_button(
                label="Download WRR Document",
                data=word_buffer,
                file_name=f"WRR_{base_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        st.markdown("---")
        st.subheader("Email Template (Copy and paste into Outlook)")
        
        email_subject = f"Warehouse Receiving Report / PO {po_number}"
        st.text_input("Subject:", value=email_subject)
        
        email_html = f"""
<div style="font-family: Calibri, Arial, sans-serif; font-size: 14px;">
<p>Dear Team,</p>
<p>Please find attached the Warehouse Receiving Report (WRR) and Excel database for PO {po_number}.<br>
Goods have been successfully received at the {wh} warehouse.<br>
The data has been extracted and transferred to Farmasoft for balance registration.</p>

<table border="1" style="border-collapse: collapse; width: 100%; max-width: 1000px; border-color: gray;">
<tr style="background-color: rgba(128, 128, 128, 0.2);">
<th style="padding: 8px; text-align: left;">PO</th>
<th style="padding: 8px; text-align: left;">Order request</th>
<th style="padding: 8px; text-align: left;">Requester</th>
<th style="padding: 8px; text-align: left;">Item Name [ENG]</th>
<th style="padding: 8px; text-align: left;">Batch</th>
<th style="padding: 8px; text-align: left;">Expiry date</th>
<th style="padding: 8px; text-align: left;">Total Quantity Inbound</th>
</tr>
<tr>
<td style="padding: 8px;">{po_number}</td>
<td style="padding: 8px;">{or_number}</td>
<td style="padding: 8px;">{requester}</td>
<td style="padding: 8px;">{item_name_eng}</td>
<td style="padding: 8px;">{batch}</td>
<td style="padding: 8px;">{exp_date}</td>
<td style="padding: 8px;">{quantity}</td>
</tr>
</table>
</div>
"""
        
        st.markdown(email_html, unsafe_allow_html=True)
        st.info("💡 Highlight the text and table above, press Ctrl+C, and paste it into your Outlook email body.")