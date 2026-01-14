import os
import io
import json
import time
import requests
import pandas as pd
import re
import openpyxl
from openpyxl.styles import Font
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# ==========================================
# [設定區]
# ==========================================
INPUT_FOLDER_ID = '1sxkmMBmgcMEPXhdhKq4PRTkFFqau5uNp'
TEMPLATE_FOLDER_ID = '1uljhJix1K9kBj_liuVfQ4cXav4zmvdNG'
PROCESSED_FOLDER_ID = '1CBUFj4ZvsSq2oWMe4kB98It-MDCXXI1J'
ARCHIVE_FOLDER_ID = '1GARx6HTKftx9r14ftZFrQzqMTbheXnZ6'
TEMPLATE_FILENAME = 'Target.xlsx'

# 【請修改這裡】貼上你剛剛在 Power Automate 產生的 HTTP POST URL
PA_EMAIL_WEBHOOK_URL = "https://default34240d49e8dc40b99241230e2426aa.f4.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0ac9b93a56384d54a7fea4a8aebf3d31/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=4zJHjwyIYEJ9uupkdJrScE1kZ7F1FTjqmLSUcfv89uw" 
# ==========================================

def get_drive_service():
    """取得 Drive 連線 (使用 token.json)"""
    creds = None
    if os.path.exists('token.json'):
        try:
            creds = Credentials.from_authorized_user_file('token.json')
        except Exception as e:
            print(f"讀取 token.json 失敗: {e}")
    elif os.environ.get('GDRIVE_TOKEN'):
        try:
            token_info = json.loads(os.environ.get('GDRIVE_TOKEN'))
            creds = Credentials.from_authorized_user_info(token_info)
        except Exception as e:
            print(f"環境變數 Token 解析失敗: {e}")

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as e:
            print(f"Token 過期且無法更新: {e}")
            return None

    if not creds:
        print("錯誤：找不到有效的 Token！")
        return None

    return build('drive', 'v3', credentials=creds)

def download_file(service, file_id, local_name):
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(local_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    # print(f"已下載: {local_name}")

def upload_file(service, local_path, folder_id, file_name):
    """
    修改過：現在會回傳 file 物件，以便取得 ID
    """
    file_metadata = {'name': file_name, 'parents': [folder_id]}
    media = MediaFileUpload(local_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    # 執行上傳並取得回傳值
    file = service.files().create(body=file_metadata, media_body=media, fields='id, name').execute()
    
    print(f"已上傳結果: {file_name} (ID: {file.get('id')})")
    return file # <--- 關鍵修改：把上傳後的檔案資訊回傳出去

def move_file_to_archive(service, file_id, file_name):
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    service.files().update(
        fileId=file_id,
        addParents=ARCHIVE_FOLDER_ID,
        removeParents=previous_parents,
        fields='id, parents'
    ).execute()
    print(f"原始檔 {file_name} 已歸檔")

def call_power_automate_webhook(file_id, file_name):
    """
    新增功能：呼叫 Power Automate 並傳送 File ID
    """
    if "http" not in PA_EMAIL_WEBHOOK_URL:
        print("⚠️ 跳過通知：未設定 PA_EMAIL_WEBHOOK_URL")
        return

    payload = {
        "file_id": file_id,          # 傳遞檔案 ID (讓 PA 去抓內容)
        "file_name": file_name,      # 傳遞檔名 (讓 PA 設定附件名稱)
        "message": "檔案已由 GitHub Action 處理完畢"
    }
    
    headers = {"Content-Type": "application/json"}

    try:
        print(f"正在通知 Power Automate 發送郵件 ({file_name})...")
        response = requests.post(PA_EMAIL_WEBHOOK_URL, data=json.dumps(payload), headers=headers)
        if response.status_code == 202 or response.status_code == 200:
            print("✅ 成功觸發 Power Automate 寄信通知！")
        else:
            print(f"⚠️ 觸發失敗，狀態碼: {response.status_code}")
            print(response.text)
    except Exception as e:
        print(f"❌ Webhook 連線錯誤: {e}")

def process_invoice(source_file, template_file, output_file):
    print(f"正在分析: {source_file} ...")
    try:
        # --- Reader Logic (與原本相同) ---
        try:
            if source_file.lower().endswith('.xls'):
                df = pd.read_excel(source_file, header=None, dtype=str, engine='xlrd')
            else:
                df = pd.read_excel(source_file, header=None, dtype=str, engine='openpyxl')
        except:
            df = pd.read_excel(source_file, header=None, dtype=str)
        
        df = df.fillna("")
        extracted_data = []

        po_pattern = re.compile(r'450\d{6,}')
        pn_strict_pattern = re.compile(r'^(SRP|CGA|CVH|CG|BVRA|MLVS|WIP|BVR|EGA|MIP|CI|ACA|SF)[A-Z0-9\.-]+', re.IGNORECASE)
        pn_fallback_pattern = re.compile(r'^(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9\.-]{5,}$')

        for idx, row in df.iterrows():
            row_values = [str(x).strip() for x in row.values]
            
            # Item No check
            item_no_val = None
            try:
                candidate = row_values[1]
                if candidate.replace('.', '', 1).isdigit() and float(candidate) > 0:
                    item_no_val = float(candidate)
            except: pass
            if item_no_val is None: continue

            found_po = ""
            found_pn = ""
            numeric_candidates = []

            for col_idx, cell in enumerate(row_values):
                if po_pattern.search(cell): found_po = cell
                elif pn_strict_pattern.search(cell) and not po_pattern.search(cell):
                    if not found_pn: found_pn = cell
                
                if col_idx != 1 and cell.replace('.', '', 1).isdigit():
                    try:
                        val = float(cell)
                        if not str(int(val)).startswith("45013"): 
                            numeric_candidates.append(val)
                    except: pass

            if not found_pn:
                for col_idx, cell in enumerate(row_values):
                    if (pn_fallback_pattern.search(cell) and not po_pattern.search(cell) and 
                        col_idx != 1 and "INVOICE" not in cell.upper()):
                        found_pn = cell
                        break

            found_qty = 0.0
            found_price = 0.0
            if numeric_candidates:
                max_val = max(numeric_candidates)
                min_val = min(numeric_candidates)
                if max_val > 0: found_qty = max_val
                if min_val > 0 and min_val != max_val: found_price = min_val

            if found_po:
                parts = found_po.split('-')
                po_final = parts[0].strip()
                line_final = parts[1].strip() if len(parts) > 1 else ""
                
                extracted_data.append([
                    po_final, line_final, found_pn, "", found_qty, "USD", found_price
                ])

        if not extracted_data: return False

        # --- Writer Logic (與原本相同) ---
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active 
        start_row = 2
        start_col = 13
        custom_font = Font(name='Calibri Light', size=9)

        for r_idx, row_data in enumerate(extracted_data):
            for c_idx, value in enumerate(row_data):
                target_cell = ws.cell(row=start_row + r_idx, column=start_col + c_idx)
                final_value = value
                if c_idx in [0, 1]: 
                    try:
                        if value and str(value).strip():
                            final_value = int(float(value))
                    except: pass
                
                target_cell.value = final_value
                target_cell.font = custom_font
                if c_idx == 4: target_cell.number_format = '#,##0'

        wb.save(output_file)
        return True

    except Exception as e:
        print(f"邏輯錯誤: {e}")
        return False

def main():
    service = get_drive_service()
    if not service: return

    # 1. 搜尋 Input
    results = service.files().list(
        q=f"'{INPUT_FOLDER_ID}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false",
        fields="files(id, name)").execute()
    files = results.get('files', [])

    if not files:
        print("Input 資料夾是空的。")
        return

    print(f"發現 {len(files)} 個檔案，開始作業...")

    # 2. 下載模板
    tpl_results = service.files().list(
        q=f"'{TEMPLATE_FOLDER_ID}' in parents and name = '{TEMPLATE_FILENAME}' and trashed = false",
        fields="files(id, name)").execute()
    
    if not tpl_results.get('files'):
        print(f"錯誤：找不到模板 {TEMPLATE_FILENAME}")
        return
    
    download_file(service, tpl_results.get('files')[0]['id'], "temp_template.xlsx")

    # 3. 逐一處理
    for file in files:
        file_id = file['id']
        file_name = file['name']
        local_in = f"temp_{file_name}"
        local_out = f"Result_{os.path.splitext(file_name)[0]}.xlsx"

        try:
            download_file(service, file_id, local_in)
            
            # 複製一份模板來工作
            import shutil
            shutil.copy("temp_template.xlsx", "work.xlsx")
            
            if process_invoice(local_in, "work.xlsx", local_out):
                # 【修改重點】: upload_file 現在會回傳上傳後的物件
                uploaded_file = upload_file(service, local_out, PROCESSED_FOLDER_ID, local_out)
                
                # 歸檔原始文件
                move_file_to_archive(service, file_id, file_name)
                
                # 【修改重點】: 呼叫 Power Automate 寄信 (帶入 file ID)
                if uploaded_file and uploaded_file.get('id'):
                    call_power_automate_webhook(uploaded_file.get('id'), local_out)
                
                print(f"檔案 {file_name} 處理成功！")
            else:
                print(f"檔案 {file_name} 無有效資料，跳過。")

        except Exception as e:
            print(f"處理失敗: {e}")
        
        finally:
            for f in [local_in, local_out, "work.xlsx"]:
                if os.path.exists(f): os.remove(f)

    if os.path.exists("temp_template.xlsx"): os.remove("temp_template.xlsx")

if __name__ == "__main__":
    main()