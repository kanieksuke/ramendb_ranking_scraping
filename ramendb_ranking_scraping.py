import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ファイルパス設定
EXCEL_PATH = "C:/Users/kanie/OneDrive/ramendb_scraping.xlsx"
CSV_PATH = "C:/Users/kanie/OneDrive/ramendb_temp.csv"

# ログイン情報
EMAIL = "kanieksuke@yahoo.co.jp"
PASSWORD = "chimpo"

# Chromeオプション設定（アクセス制限回避）
chrome_options = Options()
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # ボット検出回避
# chrome_options.add_argument("--incognito")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")  # User-Agent偽装

# ChromeDriverの起動
service = Service()  # ChromeDriverのデフォルトパスを使用
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.set_page_load_timeout(10)

def sleep():
    time.sleep(random.uniform(3, 5))
    # time.sleep(1)

# ラーメンデータベースのトップページにアクセス
url = 'https://ramendb.supleks.jp/'
print(f'Accessing: {url}')
driver.get(url)
sleep()

# ログインページへ遷移
try:
    login_link = driver.find_element(By.LINK_TEXT, "ログイン")
    login_link.click()
    print("Navigated to login page.")
    sleep()  # ページ遷移後の待機
except NoSuchElementException:
    print("Login link not found. Exiting.")
    driver.quit()
    exit()

# メールアドレスとパスワードを入力
try:
    email_input = driver.find_element(By.NAME, "mail")
    email_input.send_keys(EMAIL)
    
    password_input = driver.find_element(By.NAME, "password")
    password_input.send_keys(PASSWORD)
    
    print("Entered login credentials.")

    # ログインボタンをクリック
    login_button = driver.find_element(By.CLASS_NAME, "formbtn")
    login_button.click()
    
    print("Clicked login button.")
    sleep()  # ログイン処理の待機
except NoSuchElementException:
    print("Login form not found. Exiting.")
    driver.quit()
    exit()

print("Login successful")

# 東京版ラーメンデータベースへのリンクをクリック
try:
    ramen_link = driver.find_element(By.XPATH, "//a[@name='ramendb' and contains(@href, 'tokyo-ramendb.supleks.jp')]")
    driver.execute_script("arguments[0].scrollIntoView();", ramen_link)
    driver.execute_script("arguments[0].click();", ramen_link)
    print("Clicked Tokyo_ramen_db link. Waiting for next page.")
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "全国版"))
    )
    print("Navigated to Tokyo Ramen Database page.")
except NoSuchElementException:
    print("Link not found. Exiting.")

# ラーメンデータベースへのリンクをクリック
# input("ブラウザを手動で操作した後、Enter を押してください...") 
try:
    nationwide_link = driver.find_element(By.LINK_TEXT, "全国版")
    driver.execute_script("arguments[0].scrollIntoView();", nationwide_link)
    driver.execute_script("arguments[0].click();", nationwide_link)
    print("Clicked Nationwide Version Ramen Database page.")
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "ニューオープンのラーメン屋さんをもっと見る"))
    )
    print("Navigated to Nationwide Version Ramen Database page.")
except NoSuchElementException:
    print("Link not found. Exiting.")

driver.quit()

