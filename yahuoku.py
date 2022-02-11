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
options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
options.add_argument('--user-data-dir=' + profile_path)
options.add_argument('--lang=ja')
options.add_argument("--remote-debugging-port=9222") 

class Yahuoku():
    def item_id_get(self):
        
        driver = webdriver.Chrome(ChromeDriverManager().install() , options = options)
        wait = WebDriverWait(driver=driver, timeout=30)

        # 要素が見つかるまで、最大10秒間待機する
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
        # 出品した商品ページから商品IDを取得する関数
        time.sleep(3)
        line_num = int(1)
        for _ in range(3):
            try:  
                

                for _ in range(120):
                    # time.sleep(random.randint(180, 240))
                    driver.get("https://auctions.yahoo.co.jp/jp/show/submit?category=0")
                    time.sleep(5)
                    driver.refresh()
                    wait.until(EC.presence_of_all_elements_located)
                    time.sleep(5)
                    up = ''
                    line_num = line_num + 1
                    #行番号と列番号を指定してセルの値を取得する（左：行番号、右：列番号）
                    image_num = int(worksheet.cell(line_num, 12).value)
                    image_num_first = int(worksheet.cell(line_num, 12).value)
                    image_num_Last = int(worksheet.cell(line_num, 13).value)
                    #画像の数だけ回す
                    for elem in range(image_num_Last - image_num_first + 1):            
                        up = "C:/Users/saita/workspace/lec_rpa/image/img" + str(elem + image_num_first) + ".jpg" 
                        time.sleep(1)
                        imgup = driver.find_element_by_id('selectFile')
                        imgup.send_keys(up)
                        

                    print(up)
                    time.sleep(3)
                    
                    #商品名
                    elem = driver.find_element_by_xpath('//*[@id="fleaTitleForm"]')
                    product_name1 = worksheet.cell(line_num , 1).value
                    if product_name1 == "finish" :
                        #強制終了
                        sys.exit(1)
                    else:
                        elem.send_keys(product_name1)
                        time.sleep(1)

                    
                    #カテゴリ
                    #性別を確認
                    sex = int(worksheet.cell(line_num , 14).value)
                    print(sex)
                    #男の時
                    if sex == 1:
                        elem = driver.find_element_by_xpath('//*[@id="acMdCateChange"]')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(1)
                        elem = driver.find_element_by_xpath('//*[@id="CategorySelect"]/ul/li[1]/a/div')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(2)
                        elem = driver.find_element_by_xpath('//*[@id="23140"]/a')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(1)
                        elem = driver.find_element_by_xpath('//*[@id="23264"]/a')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(1)
                        elem = driver.find_element_by_xpath('//*[@id="2084050111"]/a')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(1)
                        elem = driver.find_element_by_xpath('//*[@id="2084053684"]/a')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(1)
                        elem = driver.find_element_by_xpath('//*[@id="updateCategory"]')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(1)
                
                    elif sex == 0:
                        elem = driver.find_element_by_xpath('//*[@id="acMdCateChange"]')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(2)
                        elem = driver.find_element_by_xpath('//*[@id="CategorySelect"]/ul/li[1]/a/div')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(2)
                        elem = driver.find_element_by_xpath('//*[@id="23140"]/a')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(2)
                        elem = driver.find_element_by_xpath('//*[@id="23268"]')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(2)
                        elem = driver.find_element_by_xpath('//*[@id="2084050117"]')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(2)
                        elem = driver.find_element_by_xpath('//*[@id="2084053703"]/a')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(2)
                        elem = driver.find_element_by_xpath('//*[@id="updateCategory"]')
                        elem.click()
                        wait.until(EC.presence_of_all_elements_located)
                        time.sleep(2)
                    #返品
            #      driver.find_element_by_xpath('//*[@id="js-PCNonPreReutnPolicyArea"]/label/span[1]').click()
                    #説明
                    
                    #商品説明
                    iframe = driver.find_element_by_id('rteEditorComposition0')
                    driver.switch_to.frame(iframe)
                    driver.find_element_by_id('0').send_keys(worksheet.cell(line_num , 3).value)
                    driver.switch_to.default_content()

                    #終了日時
                    Select(driver.find_element_by_id("ClosingYMD")).select_by_index("7")

                    #金額
                    elem = driver.find_element_by_xpath('//*[@id="auc_BidOrBuyPrice_buynow"]')
                    money = (worksheet.cell(line_num , 2) ).value
                    # money1 = int(money) + 700
                    money1 = int(money) 
                    elem.send_keys(str(money1))

                    #確認するボタン
                    elem = driver.find_element_by_xpath('//*[@id="submit_form_btn"]')
                    elem.click()
                    wait.until(EC.presence_of_all_elements_located)
                    time.sleep(1)

                    #出品するボタン
                    elem = driver.find_element_by_xpath('//*[@id="auc_preview_submit_up"]')
                    elem.click()
                    wait.until(EC.presence_of_all_elements_located)
                    time.sleep(1)
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
    yahuoku = Yahuoku()        
    yahuoku.item_id_get()
