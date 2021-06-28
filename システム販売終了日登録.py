import pyautogui as pg
import os, tkinter, tkinter.filedialog, tkinter.messagebox
import pandas as pd
import csv
import time



root = tkinter.Tk()
root.withdraw()

# CSVファイル指定   
fTyp = [("","*.csv")]
# ファイル選択ダイアログの表示
iDir = os.path.abspath(os.path.dirname(__file__))
tkinter.messagebox.showinfo("販売終了日自動登録","登録用ファイルを選択してください！")
file = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
#CSVファイルの読み込み（リスト型）
data = pd.read_csv(file).values.tolist()

#システム登録画面を見えるようにしておく（分類のタブに特に注意）

#商品コード欄をクリック
x, y = pg.locateCenterOnScreen("X:\\通箱\\★東京営業所\\西澤\\商品登録.PNG")
pg.click(x, y)
time.sleep(1)

#販売終了日（年）枠の座標を変数へ入れる
a, b = pg.locateCenterOnScreen("X:\\通箱\\★東京営業所\\西澤\\販売終了日.png")


for index in range(len(data)):
    #10アイテム登録ごとに、スクリーンショットをとって保存。（予期せぬ場合の記録）
    if  index == 0 :
        pass
    elif index % 10 == 0 :
        pg.hotkey("winleft","printscreen")
        time.sleep(2)
    else :
        pass

    
    #商品コードを入力、エンターを押す
    pg.typewrite(data[index][0])
    time.sleep(1)
    pg.press("enter")
    #反応が遅いので、15秒待機
    time.sleep(23)
    #販売終了日（年）をクリック
    pg.click(a, b)
    time.sleep(3)
   
    #販売終了日（年）を入力、エンターを押す
    pg.typewrite(str(data[index][1]))
    time.sleep(1)
    pg.press("enter")
    time.sleep(1)
    #販売終了日（月）を入力、エンターを押す
    pg.typewrite(str(data[index][2]))
    time.sleep(1)
    pg.press("enter")
    time.sleep(1)
    #販売終了日（日）を入力、エンターを押す
    pg.typewrite(str(data[index][3]))
    time.sleep(1)
    pg.press("enter")
    time.sleep(1)
    #ｆ1キーで登録
    pg.press("f1")
    #反応が遅いので、3秒待機
    time.sleep(3)
    #「保存しますか？」で「はい」の為のエンターキー
    pg.press("enter")
    #反応が遅いので、5秒待機
    time.sleep(5)








