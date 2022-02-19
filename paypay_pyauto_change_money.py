import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import os
from PIL import Image
import sys
from datetime import datetime, date, timedelta
import datetime
import pyautogui as pag
import pyocr
import pyocr.builders
import cv2
import pyperclip #クリップボードへのコピーで使用 
import tkinter

#今日
now = '{0:%Y%m%d}'.format(datetime.datetime.now())
#Pandas.dfの準備
##

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
##Tesseractのpath
path='C:\\Program Files\\Tesseract-OCR'
os.environ['PATH'] = os.environ['PATH'] + path
##global関数##
global startline
global pricecut

class PaypayChangeMoney :
    x_pre,y_pre,w_pre,h_pre = 0,0,0,0

    #textbox
    def textbox(self):
        def btn_event():
            global startline
            global pricecut
            startline = int(txt1.get())
            pricecut  = int(txt2.get())            
            root.destroy()            

        # Tkクラス生成
        root = tkinter.Tk()
        # 画面サイズ
        root.geometry('300x200')
        # 画面タイトル
        root.title('テキストボックス')
        # ラベル
        lbl1 = tkinter.Label(text='開始する行')
        lbl1.place(x=10, y=50)
        lbl2 = tkinter.Label(text='値下げする値段(円)')
        lbl2.place(x=10, y=90)
        # テキストボックス
        txt1 = tkinter.Entry(width=20)
        txt1.place(x=130, y=50)
        txt2 = tkinter.Entry(width=20)
        txt2.place(x=130, y=90)
        # ボタン
        btn = tkinter.Button(root, text='OK', command=btn_event)
        btn.place(x=140, y=170)
        # 表示
        root.mainloop()

    #スプレッドシートから必要な情報を取得する
    def getinformetion(self,line_num):
        product_name = worksheet.cell(line_num , 1).value #商品名
        description = worksheet.cell(line_num , 3).value #商品説明
        money = worksheet.cell(line_num , 2).value #金額
        image_num_first = int(worksheet.cell(line_num, 12).value)
        image_num_Last = int(worksheet.cell(line_num, 13).value)
        current_price = int(worksheet.cell(line_num, 15).value) #現在の価格
        serialno = int(worksheet.cell(line_num, 16).value) #現在の価格
        return serialno,product_name,description,money,image_num_first,image_num_Last,current_price

    
    #ポジショニング
    def posi(self,imagename):      
        print(imagename)
        for count in range(70):
            try:
                conf = (int(100) - count*0.1)*0.01 
                print(conf)                
                x,y,w,h = pag.locateOnScreen('C:/Users/saita/workspace/lec_rpa/paypay/' + imagename + '.jpg',grayscale=True,confidence=conf)
                self.x_pre,self.y_pre,self.w_pre,self.h_pre = x,y,w,h
                print(x,y,w,h)
                break

            except:
                if count == 69 :
                    raise
                else:
                    pass 
        return x,y,w,h

    #マウス移動、クリック
    def move(self,x, y, w, h):
        center_x = x + w/2
        center_y = y + h/2
        pag.moveTo(center_x, center_y)
        pag.click()
        time.sleep(3)
    
    #画面の文字認識
    def character_recognition(self,product_name) :
        #対象の画像が見つかるまでスクロールする
        # for _ in range(30):
        #     try:
                #左画面をスクリーンショット
        sc = pag.screenshot(region=(0, 0, 1919,1079)) #始点x,y、幅、高さ
        lang = 'jpn'
        sc.save('./img/jpn.png')        
        img_path = './img/{}.png'.format(lang)
        img = Image.open(img_path)
        out_path = './img/{}_{}.png'
        tools = pyocr.get_available_tools()
        tool = tools[0]

        Line_boxes = tool.image_to_string(
            img,
            lang=lang,
            builder=pyocr.builders.TextBuilder(tesseract_layout=6)
        )
        out = cv2.imread(img_path)
        print(str(Line_boxes))
        for d in Line_boxes:
            print(d.content)
            print(d.position)
            cv2.rectangle(out, d.position[0], d.position[1], (0, 0, 255), 2) #d.position[0]は認識した文字の左上の座標,[1]は右下
            cv2.imwrite(out_path.format(lang, 'Line_boxes'), out)  
            if(d.content== product_name): #Anacondaのアイコンを認識したらクリックする
                x1,y1 = d.position()[0]
                x2,y2 = d.position()[1]
        x,y,w,h = x1,y1,x2-x1,y2-y1   
        return x,y,w,h 

    
def main():
    global startline
    global pricecut
    listing = PaypayChangeMoney()
    listing.textbox()
    print(startline)
    print(pricecut)
    line_num = startline - 1
    ##paypay起動
    x, y, w, h = listing.posi("paypay")
    listing.move(x, y, w, h)
    x, y, w, h = listing.posi("mypage")
    listing.move(x, y, w, h)
    x, y, w, h = listing.posi("shuppinsitashouhin")
    listing.move(x, y, w, h)
    ##スクロールして一番下の画面に行く
    for _ in range(20):
        pag.scroll(-100)
        time.sleep(0.5)
    time.sleep(1)

    for editcount in range(120):
##画像up##
        imagenum = ''
        print("start")
        line_num = line_num + 1
        serialno,product_name,description,money,image_num_first,image_num_Last,current_price = listing.getinformetion(line_num)
        newprice = current_price - pricecut
        #出品する商品がない場合、強制終了
        if product_name == "finish" :
            sys.exit(1)
        else:
            pass
        
        ##値下げする商品を見つける##
        for elem in range(30):
            try:
                imagename = "sc" + str(serialno)
                ##new
                x, y, w, h = listing.posi(imagename) 
                print(imagename)
                break       
            except :
                ##画像が見つからない場合、スクロールする
                pag.moveTo(900, 600)
                pag.scroll(10)
                time.sleep(1)
                print("スクロール")            
        listing.move(x, y, w, h)
                
        ##編集する##
        x, y, w, h = listing.posi("henshuusuru")
        listing.move(x, y , w, h)
        time.sleep(2)
        ##スクロール##
        for _ in range(3):
            pag.scroll(-100)
            time.sleep(1)
        ##販売価格##
        x,y,w,h = pag.locateOnScreen('C:/Users/saita/workspace/lec_rpa/paypay/' + "kakakuhenkou" + '.jpg',grayscale=True,confidence=.7)
        ##画像の下端中央をクリック##
        listing.move(x, y+50, w, h)
        for _ in range(4):
            pag.hotkey('right')
            pag.hotkey('right')
            pag.hotkey('backspace')
        
        pyperclip.copy(newprice)
        ##スプレッドシートに書き込む##
        worksheet.update_cell(line_num, 15, str(newprice))
        pag.hotkey('ctrl', 'v')
        time.sleep(1)
        ##スクロール##
        for _ in range(3):
            pag.scroll(-100)
            time.sleep(1)
        ##変更する##
        x, y, w, h = listing.posi("henkousuru")
        listing.move(x, y, w, h)
        time.sleep(3)

        try:
            imagename = "brand"
            x,y,w,h = pag.locateOnScreen('C:/Users/saita/workspace/lec_rpa/paypay/' + imagename + '.jpg',grayscale=True,confidence=0.95)
            listing.move(x, y, w, h)
            time.sleep(5)
            for _ in range(3):
                pag.scroll(-100)
                time.sleep(1)        
            x, y, w, h = listing.posi("modoru_2")
            listing.move(x, y, w, h)
        except:
            for _ in range(3):
                pag.scroll(-100)
                time.sleep(1)
                ##戻る##
            x, y, w, h = listing.posi("modoru_2")
            listing.move(x, y, w, h)   
           
# main()
        
 