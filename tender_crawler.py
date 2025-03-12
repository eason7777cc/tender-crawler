from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import docx
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os

# 搜索關鍵詞列表
keywords = ['拍攝', '影片', '紀錄片', '平面影像', '影音', '影片製作', 
           '素材製作', '短片', '攝影', '錄影', '錄製']

base_url = 'https://web.pcc.gov.tw/prkms/tender/common/basic/readTenderBasic'

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # 無頭模式，不顯示瀏覽器
    chrome_options.add_argument('--no-sandbox')  # GitHub Actions 環境需要
    chrome_options.add_argument('--disable-dev-shm-usage')  # GitHub Actions 環境需要
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def search_tender(keyword):
    driver = setup_driver()
    driver.get(base_url)
    
    try:
        # 等待頁面加載
        wait = WebDriverWait(driver, 10)
        
        # 輸入標案名稱
        tender_name = wait.until(EC.presence_of_element_located((By.ID, 'tenderName')))
        tender_name.send_keys(keyword)
        
        # 選擇等標期內
        tender_status = wait.until(EC.presence_of_element_located((By.ID, 'tenderStatus')))
        tender_status.send_keys('等標期內')
        
        # 點擊搜尋按鈕
        search_button = wait.until(EC.element_to_be_clickable((By.ID, 'searchButton')))
        search_button.click()
        
        # 等待搜尋結果
        time.sleep(3)
        
        # 獲取搜尋結果
        results = driver.find_elements(By.CLASS_NAME, 'tender-result-row')
        result_list = []
        
        for result in results:
            title = result.find_element(By.CLASS_NAME, 'tender-title').text
            date = result.find_element(By.CLASS_NAME, 'tender-date').text
            result_list.append({'title': title, 'date': date})
            
        return result_list
        
    except Exception as e:
        print(f"搜尋 {keyword} 時出錯: {str(e)}")
        return []
        
    finally:
        driver.quit()

def create_word_document(results):
    doc = docx.Document()
    doc.add_heading('標案搜尋結果', 0)
    doc.add_paragraph(f'生成日期: {datetime.now().strftime("%Y-%m-%d")}')
    
    for result in results:
        doc.add_paragraph(f"標案名稱: {result['title']}")
        doc.add_paragraph(f"公告日期: {result['date']}")
        doc.add_paragraph("-" * 50)
    
    filename = f'tender_results_{datetime.now().strftime("%Y%m%d")}.docx'
    doc.save(filename)
    return filename

def send_email(filename):
    sender_email = os.environ.get('SENDER_EMAIL')  # 從環境變數獲取
    sender_password = os.environ.get('SENDER_PASSWORD')  # 從環境變數獲取
    receiver_email = "haap0716@gmail.com"
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f'每日標案搜尋結果 - {datetime.now().strftime("%Y-%m-%d")}'
    
    with open(filename, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(filename))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(filename)}"'
        msg.attach(part)
    
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print("郵件發送成功")
    except Exception as e:
        print(f"郵件發送失敗: {str(e)}")

def main():
    # 收集所有搜尋結果
    all_results = []
    for keyword in keywords:
        results = search_tender(keyword)
        all_results.extend(results)
    
    # 移除重複結果
    unique_results = []
    seen_titles = set()
    for result in all_results:
        if result['title'] not in seen_titles:
            unique_results.append(result)
            seen_titles.add(result['title'])
    
    # 生成並發送文件
    if unique_results:
        filename = create_word_document(unique_results)
        send_email(filename)
        os.remove(filename)
    else:
        print("沒有找到任何結果")

if __name__ == "__main__":
    main()
