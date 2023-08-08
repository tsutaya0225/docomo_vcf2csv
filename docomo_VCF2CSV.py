from distutils import extension
from operator import length_hint
from ssl import VERIFY_CRL_CHECK_LEAF
import tkinter as tk
from tkinter import messagebox as msgbox
from tkinterdnd2 import *
import sys, os

#VCFファイルをドロップする場所を作成
def make_window():
    root=TkinterDnD.Tk() #ウィンドウ作成
    root.title("docomo VCF to CSV") #ウィンドウタイトルを設定
    root.geometry("250x200") #ウィンドウサイズを設定
    root.resizable(width=False, height=False) #ウィンドウサイズ変更不可
    root.drop_target_register(DND_FILES) #ファイルをドロップする領域に設定
    root.dnd_bind("<<Drop>>", file_dropped) #ファイルがドロップされた時の処理

    #ラベルを表示
    label=tk.Label(root, text="VCFファイルを\nここにドロップしてください")
    label.pack(anchor=tk.CENTER, pady=50)

    #終了ボタンを表示
    button=tk.Button(root, width=10, height=2, text="終了", command=exit_clicked)
    button.pack(side=tk.BOTTOM, anchor=tk.SE, padx=10, pady=10)

    root.mainloop()

#ドロップされたファイルの処理
def file_dropped(event):
    path_str=event.data #ドロップされたファイルの文字列
    if os.path.isdir(path_str): #ディレクトリの場合
        msgbox.showinfo("docomo VCF to CSV", os.path.basename(path_str)+" はVCFファイルではありません。")
    elif not os.path.splitext(path_str.lower())[1]==".vcf":
        msgbox.showinfo("docomo VCF to CSV", os.path.basename(path_str)+" はVCFファイルではありません。")
    else:
        new_path_str=os.path.splitext(path_str)[0]+".csv"
        if(os.path.exists(new_path_str)):
            press=msgbox.askyesno("docomo VCF to CSV", os.path.basename(new_path_str)+" は既に存在します。\n上書きしますか？")
            if press=="no":
                msgbox.showinfo("docomo VCF to CSV", path_str+"は処理しませんでした。")
                return
            fin=open(path_str, "r")
            fout=open(new_path_str, 'w')
            lines=fin.readlines()
            for i in lines:
                fout.write("\""+lines[i]+"\",")
            fin.close()
            fout.close()

#終了処理
def exit_clicked():
    sys.exit()

#開始処理
if __name__ == '__main__':
    make_window()
