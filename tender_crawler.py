from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import docx
from datetime import datetime
import os
import requests

keywords = ['拍攝', '影片', '紀錄片', '平面影像', '影音', '影片製作', 
           '素材製作', '短片', '攝影', '錄影', '錄製']

base_url = 'https://web.pcc.gov.tw/prkms/tender/common/basic/readTenderBasic'

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    try:
        driver = webdriver.Chrome(options=chrome_options)
        print("ChromeDriver 初始化成功")
        return driver
    except Exception as e:
        print(f"ChromeDriver 初始化失敗: {str(e)}")
        raise

def search_tender(keyword):
    driver = setup_driver()
    driver.get(base_url)
    
    try:
        wait = WebDriverWait(driver, 10)
        print(f"開始搜尋關鍵詞: {keyword}")
        
        tender_name = wait.until(EC.presence_of_element_located((By.ID, 'tenderName')))
        tender_name.send_keys(keyword)
        print("已輸入標案名稱")
        
        tender_status = wait.until(EC.presence_of_element_located((By.ID, 'tenderStatus')))
        tender_status.send_keys('等標期內')
        print("已選擇等標期內")
        
        search_button = wait.until(EC.element_to_be_clickable((By.ID, 'searchButton')))
        search_button.click()
        print("已點擊搜尋按鈕")
        
        time.sleep(3)
        
        results = driver.find_elements(By.CLASS_NAME, 'tender-result-row')
        result_list = []
        for result in results:
            try:
                title = result.find_element(By.CLASS_NAME, 'tender-title').text
                date = result.find_element(By.CLASS_NAME, 'tender-date').text
                result_list.append({'title': title, 'date': date})
            except Exception as e:
                print(f"解析結果時出錯: {str(e)}")
        print(f"找到 {len(result_list)} 個結果")
        return result_list
        
    except Exception as e:
        print(f"搜尋 {keyword} 時出錯: {str(e)}")
        return []
        
    finally:
        driver.quit()

def create_word_document(results):
    try:
        doc = docx.Document()
        doc.add_heading('標案搜尋結果', 0)
        doc.add_paragraph(f'生成日期: {datetime.now().strftime("%Y-%m-%d")}')
        
        for result in results:
            doc.add_paragraph(f"標案名稱: {result['title']}")
            doc.add_paragraph(f"公告日期: {result['date']}")
            doc.add_paragraph("-" * 50)
        
        filename = f'tender_results_{datetime.now().strftime("%Y%m%d")}.docx'
        doc.save(filename)
        print(f"已生成文件: {filename}")
        return filename
    except Exception as e:
        print(f"生成 Word 文件失敗: {str(e)}")
        raise

def send_line_notify(message):
    token = os.environ.get('LINE_ACCESS_TOKEN')
    if not token:
        print("缺少 LINE Access Token")
        return
    
    url = "https://api.line.me/v2/bot/message/broadcast"  # 使用 broadcast API 廣播給所有好友
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    payload = {
        "messages": [{
            "type": "text",
            "text": message
        }]
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload)
        if response.status_code == 200:
            print("LINE 訊息發送成功")
        else:
            print(f"LINE 訊息發送失敗: {response.text}")
    except Exception as e:
        print(f"LINE 訊息發送錯誤: {str(e)}")

def main():
    all_results = []
    for keyword in keywords:
        results = search_tender(keyword)
        all_results.extend(results)
    
    unique_results = []
    seen_titles = set()
    for result in all_results:
        if result['title'] not in seen_titles:
            unique_results.append(result)
            seen_titles.add(result['title'])
    print(f"總共找到 {len(unique_results)} 個唯一結果")
    
    if unique_results:
        filename = create_word_document(unique_results)
        message = f"標案搜尋完成！找到 {len(unique_results)} 個結果，詳情請查看附件文件（僅限本地查看）。"
        send_line_notify(message)
        if os.path.exists(filename):
            os.remove(filename)
            print("臨時文件已刪除")
    else:
        send_line_notify("沒有找到任何標案結果。")
        print("沒有找到任何結果")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = f"程式執行失敗: {str(e)}"
        print(error_msg)
        send_line_notify(error_msg)
