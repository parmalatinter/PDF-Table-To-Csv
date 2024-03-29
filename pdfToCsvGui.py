# PDF TO CSVアプリ
# exe アプリ化コマンド
# pyinstaller -F --noconsole --name pdf_to_csv.exe --add-binary "C:\\Users\\parma\\AppData\\Local\\Programs\\Python\\Python311\\Lib\\site-packages\\tabula\\tabula-1.0.5-jar-with-dependencies.jar;./tabula/" --icon favicon.ico pdfToCsvGui.py  --add-data "favicon.ico;./"

# tabulaでコンソールが表示されるので下記を参考に対応
# https://stackoverflow.com/questions/63773517/python-tabula-script-keeps-opening-java-exe-window-how-do-i-get-it-to-use-jawaw
import glob
import os
import sys

import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import asksaveasfilename
from tkinter.ttk import Progressbar

import pandas as pd
import tabula


class PdfToCsvApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("PDF Table To CSV")
        self.geometry("800x400")

        logo = self.temp_path('favicon.ico')
        self.iconbitmap(logo)

        # ラベルとエントリーを配置するフレーム
        input_directory_frame = tk.Frame(self)
        input_directory_frame.pack(pady=10)
        input_row_count_frame = tk.Frame(self)
        input_row_count_frame.pack(pady=10)
        tk.Label(input_directory_frame, text="フォルダ:").grid(row=0, column=0)
        tk.Label(input_row_count_frame, text="読み込み開始行:").grid(row=0, column=0)
        self.input_directory_value = tk.StringVar(value='')
        self.directory_entry = tk.Entry(input_directory_frame, textvariable=self.input_directory_value, width=30)
        self.directory_entry.grid(row=0, column=1)

        # Entryウィジェットを読み取り専用にする
        self.directory_entry.config(state="readonly")
        self.header_row_count_entry = tk.Entry(input_row_count_frame, width=30)
        self.header_row_count_entry.grid(row=0, column=1)
        self.header_row_count_entry.insert(0, 1)

        # 追加ボタンを配置するフレーム
        frame = tk.Frame(self)
        frame.pack(pady=10)

        # プログレスバー
        progress_frame = tk.Frame(self)
        progress_frame.pack()

        # プログレスバー
        self.progressbar = Progressbar(progress_frame, length=400)
        self.progressbar.pack(fill="both", expand=True, pady=10)

        # logの一覧を表示するリストボックス
        self.task_list = tk.Listbox(self, bd=10)
        self.task_list.pack(fill="both", expand=True, pady=10)

        # ボタンを作成してフォルダ選択関数を呼び出す
        tk.Button(frame, text="フォルダを選択", command=self.select_folder).pack(side="left")
        tk.Button(frame, text="CSV書き出し", command=self.to_csv).pack(side="right")

        # logの一覧
        self.logs = []

    def temp_path(self, relative_path):
        try:
            # Retrieve Temp Path
            base_path = sys._MEIPASS
        except Exception:
            # Retrieve Current Path Then Error
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    # logの追加
    def add_log(self, msg):
        self.logs.append(msg)
        self.task_list.insert("end", msg)

    # logの削除
    def delete_log(self):
        self.task_list.delete(0, tk.END)

    # アプリのクラス
    def select_folder(self):
        folder_path = filedialog.askdirectory()  # フォルダ選択ダイアログを表示
        if folder_path:
            self.input_directory_value.set(folder_path)  # 変数に新しい値をセット

    def get_header_row_count_value(self):
        input_value = self.header_row_count_entry.get()
        try:
            return int(input_value)
        except ValueError:
            return 1

    def to_csv(self):
        ## 参考　https://bunkyudo.co.jp/python-tabula-01/

        base_df = None
        base_dir = self.directory_entry.get()

        self.delete_log()
        start_row_index = self.get_header_row_count_value()

        files = glob.glob(base_dir + '/*pdf')
        self.progressbar["maximum"] = len(files)

        for index, filename in enumerate(files):
            self.progressbar["value"] = index + 1
            self.update()
            self.after(1)

            dfs = tabula.read_pdf(filename, stream=True, pages='all', guess=True, lattice=False)

            if len(dfs) == 0:
                self.add_log(f'{filename} {0} 件')
                continue

            df = pd.concat(dfs, axis=0)
            count = len(df)
            self.add_log(f'{filename} {count} 件')

            base_df = pd.concat([base_df, df[start_row_index:]], axis=0)

        if base_df is not None:
            if not base_df.empty:
                filename = asksaveasfilename(filetype=[('CSV files', '*.csv')])
                if filename:
                    base_df.to_csv(filename, index=False, encoding="shift-jis")
                    self.add_log(f"{filename} にファイルを書き出しました")
        else:
            self.add_log("0件でした")


# アプリを起動する
if __name__ == "__main__":
    app = PdfToCsvApp()
    app.mainloop()
