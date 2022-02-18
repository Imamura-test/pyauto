import paypay_pyauto_change_money
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
from urllib import request
from PIL import Image
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from datetime import datetime, date, timedelta
import datetime

now = '{0:%Y%m%d}'.format(datetime.datetime.now())
#Pandas.dfの準備



profile_path = '\\Users\\saita\\AppData\\Local\\Google\\Chrome\\User Data\\seleniumpass'
options = Options()
options.add_argument('--disable-extensions')       # すべての拡張機能を無効にする。ユーザースクリプトも無効にする
options.add_argument('--proxy-server="direct://"') # Proxy経由ではなく直接接続する
options.add_argument('--proxy-bypass-list=*')      # すべてのホスト名
options.add_argument('--start-maximized')          # 起動時にウィンドウを最大化する
# options.add_argument('--incognito')          # シークレットモードの設定を付与
options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
options.add_argument('--user-data-dir=' + profile_path)
options.add_argument('--lang=ja')
options.add_argument("--remote-debugging-port=9222") 

class RakumaChangemoney:    
    def item_id_get(self):
        paypay_pyauto_change_money.PaypayChangeMoney.textbox(self)
        startline = paypay_pyauto_change_money.startline
        pricecut = paypay_pyauto_change_money.pricecut
        driver = webdriver.Chrome(ChromeDriverManager().install() , options = options)
        wait = WebDriverWait(driver=driver, timeout=30)

        # 要素が見つかるまで、最大10秒間待機する
        driver.implicitly_wait(10)

        wait.until(EC.presence_of_all_elements_located)
        print(driver.title) # ページタイトルの確認

        # Google Spread Sheetsにアクセス
        ssk = open("spread_sheet_key.txt").read()
        jf =  open("jsonf.txt").read()
        # Google Spread Sheetsにアクセス
        spread_sheet_key = str(ssk)
        jsonf = str(jf)
        profile_path = '\\Users\\saita\\AppData\\Local\\Google\\Chrome\\User Data\\seleniumpass'
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(jsonf, scope)
        gc = gspread.authorize(credentials)
        SPREADSHEET_KEY = spread_sheet_key
        worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1
        f = worksheet
        item_id = []
        driver.get('https://fril.jp/sell')
        time.sleep(3)
        driver.refresh()
        wait.until(EC.presence_of_all_elements_located)
        time.sleep(3)


        for _ in range(120):
            # 続きを見るが存在する時の処理
            try:
                elem = driver.find_element_by_id('selling-container_button')
                elem.click()
                wait.until(EC.presence_of_all_elements_located)
                time.sleep(1)

            # # 続きを見るが存在しない時の処理
            except:
                break
        posts = driver.find_elements_by_class_name('col-lg-4.col-md-4.col-sm-4.col-xs-4.text-right')
        
                
        for post in posts:
            # 商品IDの取得
            post = post.find_elements_by_class_name('btn.btn-default')[0]
            wait.until(EC.presence_of_all_elements_located)
            post = post.get_attribute("href")
            print(post)
                
            item_id.append(post)           

        
        for item in item_id:
            print("item")
            print(item)
            driver.get(item)
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(3)
            elem = driver.find_element_by_xpath('//*[@id="sell_price"]')
            price = elem.get_attribute('value')
            print(price)
            price = str(int(price) - pricecut)
            elem.clear()
            elem.send_keys(str(price))
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(3)
            elem = driver.find_element_by_xpath('//*[@id="confirm"]')
            elem.click()
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(3)
            elem = driver.find_element_by_xpath('//*[@id="submit"]')
            elem.click()
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(3)

def main():
    changemoney = RakumaChangemoney()
    changemoney.item_id_get()
