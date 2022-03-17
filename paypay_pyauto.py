
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import os
import sys
from datetime import datetime, date, timedelta
import datetime
import pyautogui as pag
from pyscreeze import ImageNotFoundException
import cv2
import pyperclip #クリップボードへのコピーで使用 

#今日
now = '{0:%Y%m%d}'.format(datetime.datetime.now())

# Google Spread Sheetsにアクセス
ssk = open("spread_sheet_key.txt").read()
jf =  open("jsonf.txt").read()
spread_sheet_key = str(ssk)
jsonf = str(jf)
profile_path = '\\Users\\saita\\AppData\\Local\\Google\\Chrome\\User Data\\seleniumpass'
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(jsonf, scope)
gc = gspread.authorize(credentials)
SPREADSHEET_KEY = spread_sheet_key
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1
f = worksheet
##Tesseractのpath
path='C:\\Program Files\\Tesseract-OCR'
os.environ['PATH'] = os.environ['PATH'] + path

class Paypay :
    x_pre,y_pre,w_pre,h_pre = 0,0,0,0

    #スプレッドシートから必要な情報を取得する
    def getinformetion(self,line_num):
        product_name = worksheet.cell(line_num , 1).value #商品名
        description = worksheet.cell(line_num , 3).value #商品説明
        money = worksheet.cell(line_num , 2).value #金額
        image_num_first = int(worksheet.cell(line_num, 12).value)
        image_num_Last = int(worksheet.cell(line_num, 13).value)
        return product_name,description,money,image_num_first,image_num_Last
    
    #ポジショニング
    def posi(self,imagename):      
        print("start1")
        for count  in range(5):
            try:
                #locateOnScreenでは左上のx座標, 左上のy座標, 幅, 高さのタプルを返す。
                print(imagename)
                
                if pag.locateOnScreen('C:/Users/saita/workspace/lec_rpa/paypay/' + imagename + '.jpg',grayscale=True,confidence=.8):
                    x,y,w,h = pag.locateOnScreen('C:/Users/saita/workspace/lec_rpa/paypay/' + imagename + '.jpg',grayscale=True,confidence=.8)
                    self.x_pre,self.y_pre,self.w_pre,self.h_pre = x,y,w,h
                    print(x,y,w,h)
                    break
            except ImageNotFoundException:
                #1秒待つ
                time.sleep(1)
        ##所定の画像が見つからない場合、ひとつ前の動作を行う。

        else:
            print(self.x_pre,self.y_pre,self.w_pre,self.h_pre) 
            self.move(self.x_pre,self.y_pre,self.w_pre,self.h_pre)
            x,y,w,h = pag.locateOnScreen('C:/Users/saita/workspace/lec_rpa/paypay/' + imagename + '.jpg',grayscale=True,confidence=.8)
            print(x,y,w,h)        
        return x,y,w,h
    
    #マウス移動、クリック
    def move(self,x, y, w, h):
        center_x = x + w/2
        center_y = y + h/2
        pag.moveTo(center_x, center_y)
        pag.click()
        time.sleep(3)
    
def main():
    listing = Paypay()
    line_num = int(1)        
    for _ in range(120):
##画像up##
        imagenum = ''
        print("start")
        line_num = line_num + 1
        product_name,description,money,image_num_first,image_num_Last = listing.getinformetion(line_num)
        #出品する商品がない場合、強制終了
        if product_name == "finish" :
            sys.exit(1)
        else:
            pass
        x, y, w, h = listing.posi("sisutemuapuri")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("media")
        listing.move(x, y, w, h)
        
        for elem in range(image_num_Last - image_num_first + 1) :
            x, y, w, h = listing.posi("import")
            listing.move(x, y, w, h)
            ##imagefileが選択されていたとき##
            try:
                x, y, w, h = listing.posi("imagefile")                   
            except :
                x, y, w, h = listing.posi("pass")
                listing.move(x, y, w, h)
                pyperclip.copy("C:/Users/saita/workspace/lec_rpa/image")
                pag.hotkey('ctrl', 'v')
                pag.hotkey('enter')
                pass        
            else:# 失敗しなかった場合は、passする
                pass           
            x, y, w, h = listing.posi("filename")
            listing.move(x, y, w, h)
            imagenum = '"' + 'img' + str(image_num_Last - elem) + '.jpg"'  
            print(imagenum)
            pyperclip.copy(imagenum)
            pag.hotkey('ctrl', 'v')
            time.sleep(0.5)
            x, y, w, h = listing.posi("hiraku")
            listing.move(x, y, w, h)        
        ##ホームに戻る##
        x, y, w, h = listing.posi("home")
        listing.move(x, y, w, h)
        ##paypay起動
        x, y, w, h = listing.posi("paypay")
        listing.move(x, y, w, h)
##ページ作成##
        ##画像UPload
        x, y, w, h = listing.posi("shuppin_plus")
        listing.move(x, y, w, h)
        ##画像選択##
        pag.moveTo(765, 195, duration=0.5)
        pag.click()
        pag.moveTo(965, 195, duration=0.5)
        pag.click()
        pag.moveTo(1165, 195, duration=0.5)
        pag.click()
        pag.moveTo(765, 395, duration=0.5)
        pag.click()
        pag.moveTo(965, 395, duration=0.5)
        pag.click()
        pag.moveTo(1165, 395, duration=0.5)
        pag.click()
        pag.moveTo(765, 595, duration=0.5)
        pag.click()       
        pag.moveTo(965, 595, duration=0.5)        
        pag.click()   
        pag.moveTo(1165, 595, duration=0.5)
        pag.click()        
            
        ##完了##
        x, y, w, h = listing.posi("kanryou")
        listing.move(x, y, w, h)
        time.sleep(3)

        ##商品名##
        x, y, w, h = listing.posi("shouhinmei")
        listing.move(x, y, w, h)
        pyperclip.copy(product_name)
        pag.hotkey('ctrl', 'v')
        time.sleep(1)
        ##カテゴリ##
        x, y, w, h = listing.posi("kategori")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("fasshion")
        listing.move(x, y, w, h)
        pag.click()
        time.sleep(3)
        x, y, w, h = listing.posi("udedokeiakuse")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("menzuudedokei")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("udedokei")
        listing.move(x, y, w, h) 
        ##商品の状態##
        try:
            ##入力中の製品はこちらですか？が出たとき
            x, y, w, h = listing.posi("kotiradesuka")
            ##画像の左端中央をクリック##
            listing.move(x, y, w-w, h)            
        except :
            pass        
        else:# 失敗しなかった場合は、passする
            pass
        x, y, w, h = listing.posi("shouhinnojoutai")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("misinyou")
        listing.move(x, y, w, h)
        ##商品説明##
        x, y, w, h = listing.posi("shouhinsetumei")
        listing.move(x, y, w, h)
        pyperclip.copy(description)
        pag.hotkey('ctrl', 'v')
        time.sleep(2)
        ##スクロール##
        for _ in range(4):
            pag.scroll(-100)
            time.sleep(3)
        ##販売価格##
        x, y, w, h = listing.posi("hanbaikakaku")
        listing.move(x, y, w, h)
        pyperclip.copy(money)
        pag.hotkey('ctrl', 'v')
        time.sleep(2)
        ##メッセージ##
        x, y, w, h = listing.posi("messeiji")
        listing.move(x, y, w, h)
        pyperclip.copy("限界まで値下げ対応いたします")
        pag.hotkey('ctrl', 'v')
        ##s出品する##
        x, y, w, h = listing.posi("shuppinsuru")
        listing.move(x, y, w, h)
        time.sleep(2)
        x, y, w, h = listing.posi("brand_shuppin")
        listing.move(x, y, w, h)
        time.sleep(2)
        x, y, w, h = listing.posi("tuduketeshuppin")
        listing.move(x, y, w, h)
        time.sleep(2)
##画像を削除##
        x, y, w, h = listing.posi("home")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("sisutemuapuri")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("media")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("export")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("subetesentaku")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("gomi")
        listing.move(x, y, w, h)
        x, y, w, h = listing.posi("home")
        listing.move(x, y, w, h)        
   
# main()
        