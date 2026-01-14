import os
from google_auth_oauthlib.flow import InstalledAppFlow

# 設定權限：允許讀寫 Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive']

def main():
    if not os.path.exists('credentials.json'):
        print("錯誤：找不到 credentials.json！")
        print("請確認你有從 Google Cloud 下載 JSON 檔，並改名放在這裡。")
        return

    print("正在開啟瀏覽器... 請登入你的 Google 帳號...")
    
    # 啟動登入流程
    flow = InstalledAppFlow.from_client_secrets_file(
        'credentials.json', SCOPES)
    
    # 取得權杖
    creds = flow.run_local_server(port=0)

    # 存檔
    with open('token.json', 'w') as token:
        token.write(creds.to_json())
        
    print("\n成功！已產生 'token.json'。")
    print("這就是你的永久通行證，程式會自動讀取它。")

if __name__ == '__main__':
    main()