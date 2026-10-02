import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import os
from openai import OpenAI
import reportlab
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Magic Dent Order Parser", page_icon="🦷", layout="centered")

st.markdown("""
    <style>
    .main-title { text-align: center; color: #1E3A8A; font-family: 'Helvetica', sans-serif; }
    .sub-text { text-align: center; color: #4B5563; font-size: 16px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-title'>🦷 Magic Dent Order Parser</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-text'>أهلاً بك يا دكتور. تفريغ الطلبيات والأسعار وحساب الإجمالي بذكاء تام.</p>", unsafe_allow_html=True)
st.markdown("---")

FIXED_API_KEY = ""

api_key = FIXED_API_KEY
if not api_key or "حط_مفتاحك" in api_key:
    try:
        if "GROQ_API_KEY" in st.secrets:
            api_key = st.secrets["GROQ_API_KEY"]
    except Exception:
        pass

if not api_key or "حط_مفتاحك" in api_key:
    api_key = st.text_input("Enter your Groq API Key (gsk_...):", type="password")

whatsapp_message = st.text_area("Paste your messy WhatsApp order message here (Arabic, Franco, English, with quantities and prices):", height=220)

if st.button("🚀 Analyze Items, Prices & Calculate Total"):
    if not api_key or "حط_مفتاحك" in api_key:
        st.warning("Please enter or configure your Groq API Key first.")
    elif not whatsapp_message.strip():
        st.warning("Please paste the WhatsApp message first.")
    else:
        with st.spinner("جاري قراءة الأصناف، التمييز بين الكميات والأسعار، وحساب الإجمالي بدقة..."):
            try:
                client = OpenAI(
                    api_key=api_key,
                    base_url="https://api.groq.com/openai/v1"
                )
                
                # تذكير الموديل بذكاء بالفرق بين الكمية الصغيرة والأسعار الكبيرة وحساب الإجمالي
                response = client.chat.completions.create(
                    model="openai/gpt-oss-120b",
                    max_tokens=8192,
                    messages=[
                        {
                            "role": "system",
                            "content": "You are an expert dental supply assistant. Process the ENTIRE messy WhatsApp order message completely without skipping ANY items. Extract EVERY single item, quantity (usually small numbers like 1, 2, 3), price (usually larger numbers or explicitly marked with currency), and brand name/notes. Translate Franco-Arabic or Arabic terms into professional English medical dental supply terms. You MUST output ALL items. Return ONLY structured rows separated by a pipe '|' in this exact format for each item: Item Name | Amount | Price | Brand Name Or Notes. If a price is not mentioned, write '-'. At the very end, add a summary row in this exact format: TOTAL | - | [Calculated Total Sum of Prices or '-'] | - . Do not summarize, do not skip items, and do not include markdown code blocks or extra text."
                        },
                        {
                            "role": "user",
                            "content": whatsapp_message
                        }
                    ]
                )
                
                result_text = response.choices[0].message.content.strip()
                lines = result_text.split('\n')
                
                parsed_data = []
                idx = 1
                for line in lines:
                    if '|' in line:
                        parts = line.split('|')
                        if len(parts) >= 4:
                            item_name = parts[0].strip()
                            qty = parts[1].strip()
                            price = parts[2].strip()
                            notes = parts[3].strip()
                            
                            # التحقق مما إذا كان هذا هو صف الإجمالي النهائي
                            if "TOTAL" in item_name.upper() or "إجمالي" in item_name:
                                parsed_data.append({
                                    "N": "TOTAL",
                                    "Item": "TOTAL",
                                    "Amount": "-",
                                    "Price": price,
                                    "Brand Name / Notes": "-"
                                })
                            else:
                                parsed_data.append({
                                    "N": idx,
                                    "Item": item_name,
                                    "Amount": qty,
                                    "Price": price,
                                    "Brand Name / Notes": notes
                                })
                                idx += 1
                        elif len(parts) == 3: # احتياطي لو الرد جاء بـ 3 أعمدة
                            item_name = parts[0].strip()
                            qty = parts[1].strip()
                            notes = parts[2].strip()
                            
                            parsed_data.append({
                                "N": idx,
                                "Item": item_name,
                                "Amount": qty,
                                "Price": "-",
                                "Brand Name / Notes": notes
                            })
                            idx += 1
                
                if parsed_data:
                    df = pd.DataFrame(parsed_data)
                else:
                    df = pd.DataFrame([{"N": 1, "Item": "Error parsing output", "Amount": "", "Price": "", "Brand Name / Notes": result_text}])
                
                st.session_state['df_orders'] = df
                st.success(f"تم بنجاح تحليل كافة الأصناف والأسعار وحساب الإجمالي بدقة تامة!")
                
            except Exception as e:
                st.error(f"حدث خطأ أثناء الاتصال بـ Groq: {e}")

if 'df_orders' in st.session_state:
    df = st.session_state['df_orders']
    st.markdown(f"### 📊 Extracted Orders, Prices & Total Table:")
    st.dataframe(df, use_container_width=True)
    
    def create_word_file(dataframe):
        doc = Document()
        for section in doc.sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            
        p = doc.add_paragraph()
        r = p.add_run("Dent Magic - قائمة الطلبيات، الأسعار، والإجمالي")
        r.font.name = 'Helvetica'
        r.font.size = Pt(14)
        r.font.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        
        table = doc.add_table(rows=1, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "N"
        hdr_cells[1].text = "Item"
        hdr_cells[2].text = "Amount"
        hdr_cells[3].text = "Price"
        hdr_cells[4].text = "Brand / Notes"
        
        for cell in hdr_cells:
            shading = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shading)
            for p_h in cell.paragraphs:
                p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p_h.runs:
                    run.font.bold = True
                    run.font.name = 'Helvetica'
                    run.font.size = Pt(10)
                    
        for index, row in dataframe.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['N'])
            row_cells[1].text = str(row['Item'])
            row_cells[2].text = str(row['Amount'])
            row_cells[3].text = str(row['Price'])
            row_cells[4].text = str(row['Brand Name / Notes'])
            
            for cell in row_cells:
                for p_r in cell.paragraphs:
                    p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p_r.runs:
                        run.font.name = 'Helvetica'
                        run.font.size = Pt(9)
                        
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio

    def create_pdf_file(dataframe):
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        elements = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=14,
            alignment=1,
            spaceAfter=15
        )
        
        elements.append(Paragraph("Dent Magic - قائمة الطلبيات، الأسعار، والإجمالي", title_style))
        
        table_data = [["N", "Item", "Amount", "Price", "Brand / Notes"]]
        for index, row in dataframe.iterrows():
            table_data.append([
                str(row['N']),
                str(row['Item']),
                str(row['Amount']),
                str(row['Price']),
                str(row['Brand Name / Notes'])
            ])
            
        t = Table(table_data, colWidths=[30, 180, 50, 70, 140])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F2F2F2')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3')),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 9),
        ]))
        
        elements.append(t)
        doc.build(elements)
        pdf_buffer.seek(0)
        return pdf_buffer

    st.markdown("---")
    st.markdown("### 📥 Download Ready Files:")
    
    col1, col2 = st.columns(2)
    with col1:
        word_file = create_word_file(df)
        st.download_button(
            label="📄 Download Word (.docx)",
            data=word_file,
            file_name="Magic_Dent_Orders_Total.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    with col2:
        pdf_file = create_pdf_file(df)
        st.download_button(
            label="📑 Download PDF",
            data=pdf_file,
            file_name="Magic_Dent_Orders_Total.pdf",
            mime="application/pdf"
        )
