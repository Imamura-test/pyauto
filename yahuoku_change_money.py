import gspread
import paypay_pyauto_change_money
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
from urllib import request
from PIL import Image
from selenium.webdriver.support.select import Select
import random
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
                 # パスを通すためのコード
from selenium.common.exceptions import TimeoutException
import sys
from datetime import datetime, date, timedelta
import datetime

#今日
now = '{0:%Y%m%d}'.format(datetime.datetime.now())
#Pandas.dfの準備

jsonf = "webscraping-7ad1c-bc2ff42a463d.json"
spread_sheet_key = "1kLMppQEqZyx8xQDyTVodsrUkze78cmbj-AqpL2UECdU"
profile_path = '\\Users\\saita\\AppData\\Local\\Google\\Chrome\\User Data\\seleniumpass'
item_not_list = open("item_not_list.txt").read().splitlines()

#chrome,Chrome Optionsの設定
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


class YahuokuChangeMoney():
    def item_id_get(self):
        paypay_pyauto_change_money.PaypayChangeMoney.textbox(self)
        startline = paypay_pyauto_change_money.startline
        pricecut = paypay_pyauto_change_money.pricecut
        item_id = [] 
        driver = webdriver.Chrome(ChromeDriverManager().install() , options = options)
        wait = WebDriverWait(driver=driver, timeout=30)
        driver.implicitly_wait(10)

        wait.until(EC.presence_of_all_elements_located)
        print(driver.title) # ページタイトルの確認

        # Google Spread Sheetsにアクセス
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(jsonf, scope)
        gc = gspread.authorize(credentials)
        SPREADSHEET_KEY = spread_sheet_key
        worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1
        f = worksheet
        next_url = str("first")

        # 出品した商品ページから商品IDを取得する関数
        time.sleep(3)        
        driver.get("https://auctions.yahoo.co.jp/openuser/jp/show/mystatus?select=selling")
        time.sleep(3)
        driver.refresh()
        wait.until(EC.presence_of_all_elements_located)
        time.sleep(5)
        try:
            while(True ):            
                for post in range(51):
                    # 商品IDの取得
                    try:
                        elem = driver.find_element_by_xpath('//*[@id="acWrContents"]/div/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[2]/td/table/tbody/tr[' + str(post + 2) + ']/td[2]/a').get_attribute('href')
                        item_id.append(elem)
                        print(elem)
                    
                    except:
                
                        next_url = driver.find_element_by_xpath('//*[@id="acWrContents"]/div/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[3]/td/table[2]/tbody/tr/td[2]/a')
                        print(next_url.text)
                        if next_url.text == "次の50件":                
                            next_url.click()
                            wait.until(EC.presence_of_all_elements_located)
                            time.sleep(3)
                            post = 0
                            

                        # 続きを見るが存在しない時の処理
                        else:
                            print("break1")
                            break
                else:
                    continue                                 

        except:
            print("item_id_getは中断されました。")
        for _ in range(3):
            try:  
            
                for id in item_id:
                    time.sleep(random.randint(10, 20))
                    driver.get(id)
                    wait.until(EC.presence_of_all_elements_located)
                    time.sleep(3)
                    elem = driver.find_element_by_xpath('//*[@id="l-sub"]/div[1]/ul/li[3]/div[2]/div/div/div[1]/a')
                    if elem.is_enabled():
                        print ("true")
                        elem.click()            
                    wait.until(EC.presence_of_all_elements_located)
                    time.sleep(3)    
                    #価格変更
                    elem = driver.find_element_by_xpath('//*[@id="BidModals"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[2]/input')
                    price = elem.get_attribute('value')
                    print(price)          
                    elem.clear()
                    price = str(int(price) - pricecut)
                    elem.send_keys(price)
                    elem = driver.find_element_by_xpath('//*[@id="BidModals"]/div[1]/div[2]/div[2]/div/div/form/div[1]/a')
                    elem.click()
                    wait.until(EC.presence_of_all_elements_located)
                    time.sleep(3)
                    elem = driver.find_element_by_xpath('//*[@id="BidModals"]/div[1]/div[2]/div[2]/div/div/form/div[2]/p/span/input')
                    elem.click()
                    wait.until(EC.presence_of_all_elements_located)
                    time.sleep(3)
            except :
                driver.save_screenshot("try_relist" + now  + ".png" )
                    # エラーメッセージを格納する
                
            else:
                    # 失敗しなかった場合は、ループを抜ける
                break
        else:
            driver.save_screenshot("error_relist" + now  + ".png" )
            sys.exit(1)

def main():
    changemoney = YahuokuChangeMoney()
    changemoney.item_id_get()