# モジュールのインポート
import os, tkinter, tkinter.filedialog, tkinter.messagebox
import openpyxl as px

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

wb = px.Workbook()
wb.save("ABCanalysis.xlsx")

