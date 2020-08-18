# モジュールのインポート
import os, tkinter, tkinter.filedialog, tkinter.messagebox
import openpyxl as px
import pandas as pd
import csv

# ファイル選択ダイアログの表示
root = tkinter.Tk()
root.withdraw()

# CSVファイル指定   
fTyp = [("","*.csv")]

iDir = os.path.abspath(os.path.dirname(__file__))
tkinter.messagebox.showinfo('ABC分析プログラム','処理ファイルを選択してください！')
file = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)

# 処理ファイル名の出力
#tkinter.messagebox.showinfo('○×プログラム',file)

#データをdfに読み込み。pandasをpdとして利用。
df = pd.read_csv(file,encoding='cp932',names=('JANcode', '商品名', '分析値'))

print(df)
#wb = px.Workbook()
#ws = wb.active
df.to_excel('ABCanalysis.xlsx',sheet_name='ABCanalysis',index=False)

#wb.save("ABCanalysis.xlsx")

