import tkinter as tk
from tkinter import messagebox as mbox
from tkinterdnd2 import *
import sys, os


# VCFファイルをドロップするウインドウを作成
def make_window():
    root = TkinterDnD.Tk()  # ウィンドウ作成
    root.title("docomo VCF to CSV")  # ウィンドウタイトルを設定
    root.geometry("250x200")  # ウィンドウサイズを設定
    root.resizable(width=False, height=False)  # ウィンドウサイズ変更不可
    root.drop_target_register(DND_FILES)  # ファイルをドロップする領域に設定
    root.dnd_bind("<<Drop>>", file_dropped)  # ファイルがドロップされた時の処理

    # ラベルを表示
    label = tk.Label(root, text="VCFファイルを\nここにドロップしてください")
    label.pack(anchor=tk.CENTER, pady=50)

    # 終了ボタンを表示
    button = tk.Button(root, width=10, height=2, text="終了", command=exit_clicked)
    button.pack(side=tk.BOTTOM, anchor=tk.SE, tmp_strpadx=10, pady=10)

    root.mainloop()


# ドロップされたファイルの処理
def file_dropped(event):
    str_path = event.data  # ドロップされたファイルの文字列
    if os.path.isdir(str_path):  # ディレクトリの場合
        mbox.showwarning("docomo VCF to CSV", os.path.basename(str_path) + " はVCFファイルではありません。")
    elif not os.path.splitext(str_path.lower())[1] == ".vcf":
        mbox.showwarning("docomo VCF to CSV", os.path.basename(str_path) + " はVCFファイルではありません。")
    else:
        str_new_path = os.path.splitext(str_path)[0] + ".CSV"
        if os.path.exists(str_new_path):
            if not mbox.askyesno(
                "docomo VCF to CSV",
                os.path.basename(str_new_path) + " は既に存在します。\n上書きしますか？",
            ):
                mbox.showinfo("docomo VCF to CSV", os.path.basename(str_path) + " は処理しませんでした。")
                return

        # 変換処理
        fin = open(str_path, "rb")  # 文字コードに依存しないようバイナリで処理
        lines = fin.readlines()
        if lines[0] != b"BEGIN:VCARD\r\n":
            mbox.showwarning("docomo VCF to CSV", os.path.basename(str_path) + " はVCF形式ではありません。")
            return
        fout = open(str_new_path, "wb")
        for line in lines:
            if line == b"X-DCM-EXPORT:manual\r\n":  # 3行目にドコモ独自の文字列があるので無視
                pass
            elif line == b"END:VCARD\r\n":
                fout.write(line)
            else:
                line = line.replace(b'"', b'""')
                line = b'"' + line
                line = line.replace(b"\r\n", b'",')
                fout.write(line)
        fin.close()
        fout.close()
        mbox.showinfo("docomo VCF to CSV", "処理を終了しました。")


# 終了処理
def exit_clicked():
    sys.exit()


# 開始処理
if __name__ == "__main__":
    make_window()
