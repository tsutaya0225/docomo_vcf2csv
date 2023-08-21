import tkinter as tk
from tkinter import messagebox as mbox
from tkinterdnd2 import *
import sys, os, csv, tempfile


# VCFファイルをドロップするウインドウを作成
def make_window():
    root = TkinterDnD.Tk()  # ウィンドウ作成
    root.title("docomo CSV to VCF")  # ウィンドウタイトルを設定
    root.geometry("250x200")  # ウィンドウサイズを設定
    root.resizable(width=False, height=False)  # ウィンドウサイズ変更不可
    root.drop_target_register(DND_FILES)  # ファイルをドロップする領域に設定
    root.dnd_bind("<<Drop>>", file_dropped)  # ファイルがドロップされた時の処理

    # ラベルを表示
    label = tk.Label(root, text="CSVファイルを\nここにドロップしてください")
    label.pack(anchor=tk.CENTER, pady=50)

    # 終了ボタンを表示
    button = tk.Button(root, width=10, height=2, text="終了", command=exit_clicked)
    button.pack(side=tk.BOTTOM, anchor=tk.SE, padx=10, pady=10)

    root.mainloop()


# ドロップされたファイルの処理
def file_dropped(event):
    str_path = event.data  # ドロップされたファイルの文字列
    if os.path.isdir(str_path):  # ディレクトリの場合
        mbox.showwarning("docomo CSV to VCF", os.path.basename(str_path) + " はCSVファイルではありません。")
    elif not os.path.splitext(str_path.lower())[1] == ".csv":
        mbox.showwarning("docomo CSV to VCF", os.path.basename(str_path) + " はCSVファイルではありません。")
    else:
        str_new_path = os.path.splitext(str_path)[0] + ".VCF"
        if os.path.exists(str_new_path):
            if not mbox.askyesno(
                "docomo CSV to VCF",
                os.path.basename(str_new_path) + " は既に存在します。\n上書きしますか？",
            ):
                mbox.showinfo("docomo CSV to VCF", os.path.basename(str_path) + " は処理しませんでした。")
                return

        # テンポラリファイルとして安全なファイル名を生成する
        fp = tempfile.NamedTemporaryFile()
        str_tmp_path = fp.name
        fp.close()

        fin = open(str_path, "r")
        ftmp = open(str_tmp_path, "w", encoding="utf-8", newline="")
        lines = fin.readlines()
        for line in lines:
            line = line.rstrip(",")
            ftmp.write(line)
        fin.close()
        ftmp.close()

        ftmp = open(str_tmp_path, "r", encoding="utf-8", newline="")
        fout = open(str_new_path, "w", encoding="utf-8")
        reader = csv.reader(ftmp)
        count = 0
        for row in reader:
            for column in row:
                if (count) == 2:
                    fout.write("X-DCM-EXPORT:manual\n")
                elif column == "":
                    pass
                else:
                    fout.write(column + "\n")
                count += 1

        fout.close()
        ftmp.close()
        os.remove(str_tmp_path)
        mbox.showinfo("docomo CSV to VCF", "処理を終了しました。")
        sys.exit()


# 終了処理
def exit_clicked():
    sys.exit()


# 開始処理
if __name__ == "__main__":
    make_window()
