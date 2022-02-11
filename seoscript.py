import paypay_pyauto_change_money
import paypay_pyauto
import yahuoku
import yahuoku_change_money
import rakuma_relist
import rakuma_change_money
import tkinter 
from tkinter import ttk

global selectsite
global selectoperation

class Seoscript:
  #textbox
    def textbox(self):

        def btn_event():
            global selectsite
            global selectoperation
            selectsite = combobox1.get()
            selectoperation = combobox2.get()            
            root.destroy()            

        # Tkクラス生成
        root = tkinter.Tk()
        # 画面サイズ
        root.geometry('400x300')
        # 画面タイトル
        root.title('入力画面')
        style = ttk.Style()
        lbl1 = tkinter.Label(text='フリマサイトを選んでください')
        lbl1.place(x=10, y=50)
        lbl2 = tkinter.Label(text='なにする？')
        lbl2.place(x=10, y=160)
        style.theme_use("winnative")
        style.configure("office.TCombobox", selectbackground="blue", fieldbackground="grey",padding=5)

        module1 = ('paypayフリマ','ヤフオク', 'ラクマ')
        v1 = tkinter.StringVar()
        combobox1 = ttk.Combobox(root, values=module1)
        combobox1 = ttk.Combobox(root, textvariable= v1, values=module1, style="office.TCombobox")
        combobox1.current(0)
        combobox1.pack(pady=10)
        combobox1.place(x=160, y=50)

        module2 = ('値下げ', '再出品')
        v2 = tkinter.StringVar()
        combobox2 = ttk.Combobox(root, values=module2)
        combobox2 = ttk.Combobox(root, textvariable= v2, values=module2, style="office.TCombobox")
        combobox2.current(0)
        combobox2.pack(pady=10)
        combobox2.place(x=160, y=160)

        # # ボタン
        btn = tkinter.Button(root, text='OK', command=btn_event)
        btn.place(x=200, y=250)
        # 表示
        root.mainloop()

def main():
    seoscript = Seoscript()
    seoscript.textbox()
    if selectsite == 'paypayフリマ' and selectoperation == '値下げ':
        paypay_pyauto_change_money.main()
    elif selectsite == 'paypayフリマ' and selectoperation == '再出品':
        paypay_pyauto.main()
    elif selectsite == 'ヤフオク' and selectoperation == '値下げ':
        yahuoku_change_money.main()
    elif selectsite == 'ヤフオク' and selectoperation == '再出品':
        yahuoku.main()
    elif selectsite == 'ラクマ' and selectoperation == '値下げ':
        rakuma_change_money.main()
    elif selectsite == 'ラクマ' and selectoperation == '再出品':
        rakuma_relist.main()

main()