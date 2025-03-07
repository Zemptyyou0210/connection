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
    "ER": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Lorazepam 2mg/mL/Amp": 5
    },
    "OR": {
        "Morphine HCl 10mg/1mL/Amp": 20
    },
    "麻醉科": {
        "MORPHINE20mg/1mL/Amp(PCA用)(5 amp/包)":10,
        "Morphine HCl 10mg/1mL/Amp":3,
        "Meperidine(Pethidine) 50mg/mL/Amp":3,
        "Fentanyl(0.05mg/mL) 2mL/Amp":400,
        "Fentanyl inj 0.05mg/mL 10mL/Amp":110,
        "Fentanyl inj 0.05mg/mL 10mL/Amp(PCA)(4 amp/包)":40,
        "Alfentanil 0.5mg/mL 2mL/Amp":100,
        #"Codeine phosphate 15mg/mL/Amp":,
        "Ketamine 500mg/10mL/Vial":4,
        #"Lorazepam 2mg/mL/Amp":,
        #"Midazolam 15mg/3mL/Amp":,
        "MIDazolam 五mg/mL/Amp":150,
        "Thiamylal 300mg/Amp":11,
        #"Diazepam 10mg/2mL/Amp":,
        "Propofol 200mg/20mL/Amp":600,
        "Etomidate 20mg/10mL/amp":15
    },
    
   "內視鏡": {
       "Fentanyl(0.05mg/mL) 2mL/Amp":30,
       "MIDazolam 五mg/mL/Amp":30
    },
    
   "胸腔科檢查室": {
       "Meperidine(Pethidine) 50mg/mL/Amp":2,
       "MIDazolam 五mg/mL/Amp":10
    },
    
    "心導管": {
        "Morphine HCl 10mg/1mL/Amp":5,
        "Midazolam 15mg/3mL/Amp":5
    },
    
    "POR": {"Morphine HCl 10mg/1mL/Amp":10,
            "Meperidine(Pethidine) 50mg/mL/Amp":3,
            "Fentanyl(0.05mg/mL) 2mL/Amp":40,
            "Midazolam 15mg/3mL/Amp":2,
            "Diazepam 10mg/2mL/Amp":1
    },

    "6A": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 3,
        "Codeine phosphate 15mg/imL INJ": 3,
        # "Lorazepam 2mg/mL/Amp": 35
    },
    
       
    "6A": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 3,
        "Codeine phosphate 15mg/imL INJ": 3,
   
    },
    "MICU": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Fentanyl inj 0.05mg/mL 10mL/Amp":60,
        "Lorazepam 2mg/mL/Amp": 2,
        "Propofol 200mg/20mL/Amp":20
        
    },
    "SICU": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 5,
        "Fentanyl inj 0.05mg/mL 10mL/Amp":20,
        "Codeine phosphate 15mg/imL INJ": 5,
        "Lorazepam 2mg/mL/Amp": 2, 
        "Midazolam 15mg/3mL/Amp":2,
        "Propofol 200mg/20mL/Amp":10
    },
    "7A": {
        "Morphine HCl 10mg/1mL/Amp": 30,
        "Meperidine(Pethidine) 50mg/mL/Amp": 3,
        "Codeine phosphate 15mg/imL INJ": 5,

    },
    "7B": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 3,

    },
    "8A": {
        "Morphine HCl 10mg/1mL/Amp": 40,
        "Meperidine(Pethidine) 50mg/mL/Amp": 2,

    },
    "8B": {
        "Morphine HCl 10mg/1mL/Amp": 20,


    },
    "9A": {
        "Morphine HCl 10mg/1mL/Amp": 20,
        "Meperidine(Pethidine) 50mg/mL/Amp": 3,
        "Codeine phosphate 15mg/imL INJ": 4,
        "Lorazepam 2mg/mL/Amp": 1
    },
    "9B": {
        "Morphine HCl 10mg/1mL/Amp": 15,

    }
       
}

# 定義列名
COLUMNS = ["現存量", "空瓶", "處方箋", "效期>6個月", "常備量=線存量+空瓶(空瓶量=處方箋量)", "備註"]

# 定義查核藥師列表
PHARMACISTS =['', '廖文佑', '洪英哲', '楊曜嘉', '劉芷妘', '郭莉萱','蔡尚憲','鍾向渝', '吳雨柔', '侯佳旻', '蘇宜萱', '王孝軒', '王奕祺', '周芷伊', '簡妙格', '陳威如', 
               '邱柏翰', '紀晨雲', '吳振凌', '羅志軒', '王威智', '劉川葆', '江廷昌','凃惠敏', '張淑娟', 
               '李典則', '熊麗婷', '許家誠', '盧柏融', '劉奕君', '張雯婷', '張亦汝', '陳意涵','林坤瑝', '蔡文子','王奕祺', '邱柏翰',] 

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
            if col == "現存量":
                drug_data[col] = st.number_input(
                    f"{col} ({drug})",
                    min_value=0,
                    max_value=limit,
                    value=limit,
                    key=f"{drug}_{col}",
                    help=f"庫存限制: {limit}支"
                )
                if drug_data[col] > limit * 0.8:  # 如果庫存超過限制的80%，顯示警告
                    st.warning(f"注意：{drug}的庫存接近或超過限制（{limit}支）")
            elif col in ["空瓶", "處方箋"]:
                drug_data[col] = st.number_input(f"{col} ({drug})", min_value=0, value=0, key=f"{drug}_{col}")
            elif col == "效期>6個月":
                expiry_status = st.selectbox(f"{col} ({drug})", ["符合", "不符合"], key=f"{drug}_{col}")
                if expiry_status == "不符合":
                    expiry_reason = st.text_area(f"不符合原因 ({drug})", key=f"{drug}_{col}_reason")
                    drug_data[col] = f"不符合: {expiry_reason}" if expiry_reason else "不符合"
                else:
                    drug_data[col] = "符合"
            elif col == "常備量=線存量+空瓶(空瓶量=處方箋量)":
                stock_status = st.selectbox(f"{col} ({drug})", ["符合", "不符合"], key=f"{drug}_{col}")
                if stock_status == "不符合":
                    stock_reason = st.text_area(f"不符合原因 ({drug})", key=f"{drug}_{col}_reason")
                    drug_data[col] = f"不符合: {stock_reason}" if stock_reason else "不符合"
                else:
                    drug_data[col] = "符合"
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
            df = pd.DataFrame(columns=['單位', '常備品項', '常備量', '現存量', '空瓶', '處方箋', '效期>6個月', '常備量=線存量+空瓶(空瓶量=處方箋量)', '日期', '單位主管', '查核藥師', '備註'])
            
            for drug, info in data.items():
                row = {
                    '單位': ward,
                    '常備品項': drug,
                    '常備量': WARD_DRUGS[ward][drug],
                    '現存量': info['現存量'],
                    '空瓶': info['空瓶'],
                    '處方箋': info['處方箋'],
                    '效期>6個月': info['效期>6個月'],
                    '常備量=線存量+空瓶(空瓶量=處方箋量)': info['常備量=線存量+空瓶(空瓶量=處方箋量)'],
                    '日期': selected_date.strftime("%Y/%m/%d"),
                    '單位主管': '',  # 這裡留空，因為簽名會單獨放在另一個工作表
                    '查核藥師': pharmacist,
                    '備註': info['備註']
                }
                df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

            # 保存為 Excel 文件
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='藥品庫存查核', index=False)
                
                # 調整列寬
                worksheet = writer.sheets['藥品庫存查核']
                for idx, col in enumerate(df.columns):
                    max_length = max(df[col].astype(str).map(len).max(), len(col))
                    worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx+1)].width = max_length + 2

                # 將簽名保存為圖片
                img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                
                # 將簽名圖片添加到新的工作表
                worksheet = writer.book.create_sheet('病房單位主管簽名')
                img = XLImage(io.BytesIO(img_byte_arr))
                worksheet.add_image(img, 'A1')

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
                revision_style = ParagraphStyle('RevisionStyle', fontName='KaiU', fontSize=9, alignment=2)  # 改回右對齊
                styles = getSampleStyleSheet()
                wrap_style = styles['Normal']  # 從預設的 'Normal' 樣式中獲取
                wrap_style.wordWrap = 'CJK'  # 支援中文自動換行
                wrap_style.leading = 10  # 設定行距
                                

                # 添加查核時間、標題和修訂日期
                check_time ="查核時間 : "+ selected_date.strftime("%Y/%m/%d")
                title_table_data = [
                    ['', Paragraph('單位庫存1-4級管制藥品月查核表', title_style), ''],
                    [Paragraph(check_time, chinese_style), '', Paragraph('114.02.27 更新', revision_style)]
                ]
                title_table = Table(title_table_data, colWidths=[page_height*0.2, page_height*0.6, page_height*0.2])
                title_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # 所有單元格居中對齊
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 10),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                    ('SPAN', (1, 0), (1, 0)),  # 標題橫跨整行
                    ('ALIGN', (0, 1), (0, 1), 'LEFT'),  # 查核時間左對齊
                    ('ALIGN', (2, 1), (2, 1), 'RIGHT'),  # 修訂時間右對齊
                ]))

                story.append(title_table)
                story.append(Spacer(1, 5*mm))  # 減少標題和表格之間的間距

                # 創建簽名圖片
                img = ReportLabImage(BytesIO(img_byte_arr))
                img.drawHeight = 15*mm
                img.drawWidth = 30*mm

                # 準備表格數據
                table_data = [
                    ['單位', '常備品項', '常備量', '查核內容', '', '', '', '', '日期', '單位主管', '查核藥師', '備註'],
                    ['', '', '', '現存量', '空瓶', '處方箋', '常備量=線存量+空瓶(空瓶量=處方箋量)', '效期>6個月', '', '', '', '']
                ]

                # 添加藥品數據
                for drug, info in data.items():

                    expiry_paragraph = Paragraph(str(info['效期>6個月'])) # 讓「效期>6個月」自動換行
                    stock_paragraph = Paragraph(str(info['常備量=線存量+空瓶(空瓶量=處方箋量)']))  # 讓「常備量=線存量+空瓶(空瓶量=處方箋量)」自動換行
                    remark_paragraph = Paragraph(str(info['備註']), wrap_style)  # 讓「備註」自動換行
                    row = [
                        ward,
                        Paragraph(drug, wrap_style),  # 藥品名稱也可以自動換行
                        str(WARD_DRUGS[ward][drug]),
                        str(info['現存量']),
                        str(info['空瓶']),
                        str(info['處方箋']),
                        expiry_paragraph,  # 自動換行的「效期>6個月」
                        stock_paragraph,  # 自動換行的「常備量=線存量+空瓶(空瓶量=處方箋量)」
                        selected_date.strftime("%Y/%m/%d"),
                        img,  # 自動換行的「單位主管」
                        pharmacist,
                        remark_paragraph  # 自動換行的「備註」
                    ]
                    table_data.append(row)

                # 創建表格，調整列寬以適應 A4 橫向
                available_width = page_height - 10*mm
                col_widths = [10*mm, 45*mm, 10*mm, 10*mm, 10*mm, 10*mm, 49*mm, 40*mm, 20*mm, 30*mm, 20*mm, 23*mm]
                table = Table(table_data, colWidths=col_widths)

                # 設置表格樣式
                table.setStyle(TableStyle([
                    ('FONT', (0, 0), (-1, -1), 'KaiU'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('BACKGROUND', (0, 0), (-1, 1), colors.lightgrey),
                
                    # 合併「查核內容」標題
                    ('SPAN', (3, 0), (7, 0)),
                    ('SPAN', (9, 2), (9, -1)),
                    # 讓這些欄位內容自動換行
                    ('ALIGN', (6, 2), (6, -1), 'LEFT'),  # 效期>6個月
                    ('ALIGN', (7, 2), (7, -1), 'LEFT'),  # 常備量=線存量+空瓶(空瓶量=處方箋量)
                  
                    ('ALIGN', (11, 2), (11, -1), 'LEFT'),  # 備註
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
