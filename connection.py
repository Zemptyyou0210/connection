import logging
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import io
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ReportLabImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from io import BytesIO
import os
import re
import traceback
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from datetime import date

# 設置日誌記錄
logging.basicConfig(level=logging.INFO)

# 定義不同病房的藥品列表和庫存限制
WARD_DRUGS = {
    "6A": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 3,
        "Codeine phosphate 15mg/imL INJ": 5,
        # "Lorazepam 2mg/mL/Amp": 35
    },
    "MICU": {
        "Morphine HCl 10mg/1mL/Amp": 30,
        "Meperidine(Pethidine) 50mg/mL/Amp": 25,
        "Codeine phosphate 15mg/imL INJ": 20,
        "Lorazepam 2mg/mL/Amp": 35
    },
    "SICU": {
        "Morphine HCl 10mg/1mL/Amp": 30,
        "Meperidine(Pethidine) 50mg/mL/Amp": 25,
        "Codeine phosphate 15mg/imL INJ": 20,
        "Lorazepam 2mg/mL/Amp": 35
    },
    "7A": {
        "Morphine HCl 10mg/1mL/Amp": 30,
        "Meperidine(Pethidine) 50mg/mL/Amp": 3,
        "Codeine phosphate 15mg/imL INJ": 5,
        # "Lorazepam 2mg/mL/Amp": 35
    },
    "7B": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 5,
        # "Codeine phosphate 15mg/imL INJ": 20,
        # "Lorazepam 2mg/mL/Amp": 35
    },
    "8A": {
        "Morphine HCl 10mg/1mL/Amp": 40,
        "Meperidine(Pethidine) 50mg/mL/Amp": 5,
        # "Codeine phosphate 15mg/imL INJ": 20,
        # "Lorazepam 2mg/mL/Amp": 35
    },
    "8B": {
        "Morphine HCl 10mg/1mL/Amp": 30,
        "Meperidine(Pethidine) 50mg/mL/Amp": 5,
        # "Codeine phosphate 15mg/imL INJ": 20,
        # "Lorazepam(針劑) 2mg/mL/Amp": 35
    },
    "9A": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 20,
        "Codeine phosphate 15mg/imL INJ": 5,
        "Lorazepam 2mg/mL/Amp": 1
    },
    "9B": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 0,
        "Codeine phosphate 15mg/imL INJ": 0,
        "Lorazepam 2mg/mL/Amp": 0
    },
        "POR": {
        "Morphine HCl 10mg/1mL/Amp": 10,
        "Meperidine(Pethidine) 50mg/mL/Amp": 3,
        "Fentanyl (0.05mg/mL) 2mL/Amp": 40,
        "Midazolam 15mg/3mL/Amp": 2,
        "Diazepam 10mg/2mL/Amp": 1
    }
}

# 定義列名
COLUMNS = ["現貨", "空瓶", "處方箋", "EXP>6month", "是否符合", "備註"]

# 定義查核藥師列表
PHARMACISTS = ["", "廖文佑", "洪英哲"]

# 設置 Google Drive API 認證
try:
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["google_drive_credentials"],
        scopes=['https://www.googleapis.com/auth/drive.file']
    )
    drive_service = build('drive', 'v3', credentials=creds)
    st.write("Debug: Google Drive API 認證成功設置")
except Exception as e:
    st.error(f"設置 Google Drive API 認證時發生錯誤: {str(e)}")
    st.exception(e)

def upload_to_drive(file_name, mime_type, file_content):
    folder_id = st.secrets["google_drive"]["folder_id"]
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaIoBaseUpload(file_content, mimetype=mime_type, resumable=True)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return file.get('id')

def create_drug_form(ward, drugs):
    data = {}
    for drug, limit in drugs.items():
        st.subheader(drug)
        drug_data = {}
        for col in COLUMNS:
            if col == "現貨":
                drug_data[col] = st.number_input(
                    f"{col} ({drug})",
                    min_value=0,
                    max_value=limit,
                    value=0,
                    key=f"{drug}_{col}",
                    help=f"庫存限制: {limit}支"
                )
                if drug_data[col] > limit * 0.8:  # 如果庫存超過限制的80%，顯示警告
                    st.warning(f"注意：{drug}的庫存接近或超過限制（{limit}支）")
            elif col in ["空瓶", "處方箋", "EXP>6month"]:
                drug_data[col] = st.number_input(f"{col} ({drug})", min_value=0, value=0, key=f"{drug}_{col}")
            elif col == "是否符合":
                drug_data[col] = st.selectbox(f"{col} ({drug})", ["Y", "N"], key=f"{drug}_{col}")
            elif col == "備註":
                drug_data[col] = st.text_area(f"{col} ({drug})", key=f"{drug}_{col}")
        data[drug] = drug_data
    return data

def main():
    st.title("藥品庫存查核表")

    # 使用 st.empty() 創建一個佔位符
    date_input_container = st.empty()

    # 獲取今天的日期
    today = date.today()

    # 使用唯一的 key 創建 date_input
    selected_date = date_input_container.date_input(
        "選擇日期",
        today,
        max_value=today,
        key="date_input_unique_key"
    )

    # 選擇病房
    ward = st.selectbox("請選擇病房", list(WARD_DRUGS.keys()))

    # 獲取該病房的藥品列表和庫存限制
    drugs = WARD_DRUGS[ward]

    # 創建藥品表單
    data = create_drug_form(ward, drugs)

    # 添加查核藥師下拉選單
    pharmacist = st.selectbox("查核藥師", PHARMACISTS, help="請選擇查核藥師")

    # 添加電子簽名畫布
    st.write("請在下方簽名：")
    st.caption("使用滑鼠或觸控筆在下方空白處簽名")
    canvas_result = st_canvas(
        fill_color="rgba(255, 165, 0, 0.3)",
        stroke_width=2,
        stroke_color="#000000",
        background_color="#ffffff",
        height=150,
        drawing_mode="freedraw",
        key="canvas",
    )

    st.write("Debug: Starting main function")
    st.write(f"Debug: upload_to_drive function exists: {'upload_to_drive' in globals()}")
    
    # 在函數開始時初始化這些變量
    excel_filename = None
    pdf_filename = None
    excel_buffer = None
    pdf_buffer = None

    if st.button("提交", key="submit_button_unique_key"):
        st.write(f"Debug: canvas_result.image_data is None: {canvas_result.image_data is None}")
        st.write(f"Debug: pharmacist: {pharmacist}")
        
        if canvas_result.image_data is None:
            st.error("請在畫布上簽名")
        elif not pharmacist:
            st.error("請選擇查核藥師")
        elif canvas_result.image_data is not None and pharmacist:
            # 使用選擇的日期
            file_date = selected_date.strftime("%Y.%m.%d")
            
            # 創建文件名（不包含副檔名）
            file_base_name = f"{file_date}_{ward}_藥品庫存查核表"
            
            # 創建 Excel 和 PDF 文件名
            excel_filename = f"{file_base_name}.xlsx"
            pdf_filename = f"{file_base_name}.pdf"

            # 創建 DataFrame
            df = pd.DataFrame(data).T
            
            # 添加查核藥師欄位
            df["查核藥師"] = pharmacist

            # 重新排序列
            columns_order = COLUMNS + ["查核藥師"]
            df = df[columns_order]

            # 保存為 Excel 文件
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='藥品庫存查核', index=True)
                
                # 將簽名保存為圖片
                img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                
                # 將簽名圖片添加到新的工作表
                worksheet = writer.book.create_sheet('病房單位主管簽名')
                img = XLImage(io.BytesIO(img_byte_arr))
                worksheet.add_image(img, 'A1')
            
            excel_buffer.seek(0)

            # 生成 PDF 文件
            pdf_buffer = io.BytesIO()
            try:
                # 創建 PDF 文檔，使用 A4 橫向
                page_width, page_height = A4
                doc = SimpleDocTemplate(pdf_buffer, pagesize=(page_height, page_width), leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
                story = []
                styles = getSampleStyleSheet()

                # 註冊字體
                pdfmetrics.registerFont(TTFont('KaiU', 'fonts/kaiu.ttf'))  # 標楷體
                pdfmetrics.registerFont(TTFont('Calibri', 'fonts/calibri.ttf'))    # Calibri

                # 創建包含中文字體的樣式
                title_style = ParagraphStyle('TitleStyle', fontName='KaiU', fontSize=16, alignment=1)
                chinese_style = ParagraphStyle('ChineseStyle', fontName='KaiU', fontSize=9)
                english_style = ParagraphStyle('EnglishStyle', fontName='Calibri', fontSize=9)
                revision_style = ParagraphStyle('RevisionStyle', fontName='KaiU', fontSize=9, alignment=2)  # 右對齊

                # 添加標題和修訂日期
                title_table_data = [
                    [Paragraph('單位庫存1-4級管制藥品月查核表', title_style), Paragraph('113.09.30 修訂', revision_style)]
                ]
                title_table = Table(title_table_data, colWidths=[page_height*0.8, page_height*0.2])
                title_table.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 0),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ]))

                story.append(title_table)
                story.append(Spacer(1, 5*mm))

                # 創建簽名圖片
                img = ReportLabImage(BytesIO(img_byte_arr))
                img.drawHeight = 15*mm
                img.drawWidth = 30*mm

                # 準備表格數據
                table_data = [
                    ['病房單位', 'DRUG', '常備量', '查核內容', '', '', '', '', '', '日期', '單位主管', '查核藥師', '備註'],
                    ['', '', '', '現貨', '空瓶', '處方箋', 'Exp>6M', '符合', '不符合', '', '', '', '']
                ]

                # 添加藥品數據
                for drug, info in data.items():
                    row = [
                        ward,
                        Paragraph(drug, english_style),
                        str(WARD_DRUGS[ward][drug]),
                        str(info['現貨']),
                        str(info['空瓶']),
                        str(info['處方箋']),
                        str(info['EXP>6month']),
                        'V' if info['是否符合'] == 'Y' else '',
                        'V' if info['是否符合'] == 'N' else '',
                        selected_date.strftime("%Y/%m/%d"),
                        img,
                        pharmacist,
                        Paragraph(info['備註'], chinese_style)
                    ]
                    table_data.append(row)

                # 創建表格，調整列寬以適應 A4 橫向
                available_width = page_height - 20*mm
                col_widths = [25*mm, 50*mm, 15*mm, 15*mm, 15*mm, 15*mm, 15*mm, 12*mm, 12*mm, 20*mm, 30*mm, 20*mm, 33*mm]
                table = Table(table_data, colWidths=col_widths)

                # 設置表格樣式
                table.setStyle(TableStyle([
                    ('FONT', (0, 0), (-1, -1), 'KaiU'),
                    ('FONT', (1, 2), (1, -1), 'Calibri'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('SPAN', (3, 0), (8, 0)),
                    ('SPAN', (10, 2), (10, -1)),  # 合併單位主管欄位
                    ('BACKGROUND', (0, 0), (-1, 1), colors.lightgrey),
                ]))

                story.append(table)

                # 生成 PDF
                doc.build(story)
                pdf_buffer.seek(0)

                st.write(f"Debug: excel_filename = {excel_filename}")
                st.write(f"Debug: pdf_filename = {pdf_filename}")
                st.write(f"Debug: excel_buffer is None: {excel_buffer is None}")
                st.write(f"Debug: pdf_buffer is None: {pdf_buffer is None}")

            except Exception as e:
                st.error(f"生成 PDF 時發生錯誤: {str(e)}")
                st.exception(e)

        # 在上傳文件之前檢查所有必要的變量是否已定義
        if excel_filename and pdf_filename and excel_buffer and pdf_buffer:
            st.write("Debug: 所有必要的變量都已設置")
            try:
                excel_file_id = upload_to_drive(excel_filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', excel_buffer)
                if excel_file_id:
                    st.success(f"Excel 文件已上傳，ID: {excel_file_id}")
                    excel_url = f"https://drive.google.com/file/d/{excel_file_id}/view"
                    st.markdown(f"[點擊此處查看 Excel 文件]({excel_url})")
                else:
                    st.error("Excel 文件上傳失敗")

                pdf_file_id = upload_to_drive(pdf_filename, 'application/pdf', pdf_buffer)
                if pdf_file_id:
                    st.success(f"PDF 文件已上傳，ID: {pdf_file_id}")
                    pdf_url = f"https://drive.google.com/file/d/{pdf_file_id}/view"
                    st.markdown(f"[點擊此處查看 PDF 文件]({pdf_url})")
                else:
                    st.error("PDF 文件上傳失敗")
            except Exception as e:
                st.error(f"上傳文件失敗: {str(e)}")
                st.exception(e)
        else:
            st.error("無法上傳文件：部分必要資訊缺失")
            st.write(f"Debug: excel_filename = {excel_filename}")
            st.write(f"Debug: pdf_filename = {pdf_filename}")
            st.write(f"Debug: excel_buffer is None: {excel_buffer is None}")
            st.write(f"Debug: pdf_buffer is None: {pdf_buffer is None}")

if __name__ == "__main__":
    main()
