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
import sys
from datetime import datetime, date, timedelta
import datetime

now = '{0:%Y%m%d}'.format(datetime.datetime.now())
#Pandas.dfの準備
profile_path = '\\Users\\saita\\AppData\\Local\\Google\\Chrome\\User Data\\seleniumpass'

#chrome,Chrome Optionsの設定
options = Options()
options.add_argument('--disable-extensions')       # すべての拡張機能を無効にする。ユーザースクリプトも無効にする
options.add_argument('--proxy-server="direct://"') # Proxy経由ではなく直接接続する
options.add_argument('--proxy-bypass-list=*')      # すべてのホスト名
options.add_argument('--start-maximized')          # 起動時にウィンドウを最大化する
options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
options.add_argument('--user-data-dir=' + profile_path)
options.add_argument('--lang=ja')
options.add_argument("--remote-debugging-port=9222") 

class RakumaRelist:
    def item_id_get(self):
        driver = webdriver.Chrome(ChromeDriverManager().install() , options = options)
        wait = WebDriverWait(driver=driver, timeout=30)
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
        line_num = int(1)       
        for _ in range(120):
            #今日

            time.sleep(1)
            driver.get("https://fril.jp/item/new")
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(2)
            up = ''
            line_num = line_num + 1
            #行番号と列番号を指定してセルの値を取得する（左：行番号、右：列番号）
            image_num_first = int(worksheet.cell(line_num, 12).value)
            image_num_Last = int(worksheet.cell(line_num, 13).value)
            for elem in range(image_num_Last - image_num_first + 1):
                print(image_num_Last)
                print(image_num_first)           
                up = "C:/Users/saita/workspace/lec_rpa/image/img" + str(elem + image_num_first) + ".jpg" 
                time.sleep(1)
                imgups = driver.find_element_by_xpath('//*[@id="files"]/div[' + str(elem + 1 )  + ']/div[1]/div/input')
                imgups.send_keys(up)
            print(up)   
            time.sleep(2)        
            #商品名
            elem = driver.find_element_by_xpath('//*[@id="name"]')
            product_name1 = worksheet.cell(line_num , 1).value
            if product_name1 == "finish" :
                #強制終了
                sys.exit(1)
            else:
                elem.send_keys(product_name1)
                time.sleep(1)

            #商品説明
            elem = driver.find_element_by_tag_name('textarea')
            elem.send_keys(worksheet.cell(line_num , 3).value)
            time.sleep(3)

            #カテゴリ
            elem = driver.find_element_by_xpath('//*[@id="category_name"]')
            elem.click()
            time.sleep(1)
            elem = driver.find_element_by_xpath('//*[@id="select-category"]/div/div/div[2]/div/div[2]/a')
            elem.click()
            time.sleep(1)
            elem = driver.find_element_by_xpath('//*[@id="menu_1"]/div[10]/a')
            elem.click()
            time.sleep(1)
            elem = driver.find_element_by_xpath('//*[@id="menu_1_9"]/a[1]')
            elem.click()
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(1)
            
            #金額
            elem = driver.find_element_by_xpath('//*[@id="sell_price"]')
            money = worksheet.cell(line_num , 2).value
            # money1 = int(money) + 700
            money1 = int(money) 
            elem.send_keys(str(money1))
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(3)

            #確認するボタン
            elem = driver.find_element_by_xpath('//*[@id="confirm"]')
            elem.click()
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(5)

            #出品するボタン
            elem = driver.find_element_by_xpath('//*[@id="submit"]')
            elem.click()
            wait.until(EC.presence_of_all_elements_located)
            time.sleep(15)

def maim():
    relist = RakumaRelist()
    relist.item_id_get()



