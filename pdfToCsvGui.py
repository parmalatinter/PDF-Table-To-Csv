# PDF TO CSVアプリ
# exe アプリ化コマンド
# pyinstaller --onefile --noconsole --name pdf_to_csv.exe --add-binary "C:\\Users\\parma\\AppData\\Local\\Programs\\Python\\Python311\\Lib\\site-packages\\tabula\\tabula-1.0.5-jar-with-dependencies.jar;./tabula/" gui.py

import glob
import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import asksaveasfilename
from tkinter.ttk import Progressbar

import pandas as pd
import tabula


class PdfToCsvApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("PDF TO CSV")
        self.geometry("800x400")

        # ラベルとエントリーを配置するフレーム
        input_directory_frame = tk.Frame(self)
        input_directory_frame.pack(pady=10)
        input_row_count_frame = tk.Frame(self)
        input_row_count_frame.pack(pady=10)
        tk.Label(input_directory_frame, text="フォルダ:").grid(row=0, column=0)
        tk.Label(input_row_count_frame, text="ヘッダー数:").grid(row=0, column=0)
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

        # tk.Button(progress_frame, text="To CSV", command=self.to_csv).pack(side="right")
        # プログレスバー
        self.progressbar = Progressbar(progress_frame, length=400)
        self.progressbar.pack(fill="both", expand=True)

        # タスクの一覧を表示するリストボックス
        self.task_list = tk.Listbox(self, bd=10)
        self.task_list.pack(fill="both", expand=True)

        # ボタンを作成してフォルダ選択関数を呼び出す
        tk.Button(frame, text="フォルダを選択", command=self.select_folder).pack(side="left")
        tk.Button(frame, text="To CSV", command=self.to_csv).pack(side="right")

        # logの一覧
        self.logs = []


    # タスクの追加
    def add_log(self, msg):
        self.logs.append(msg)
        self.task_list.insert("end", msg)

    # タスクの削除
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
            numeric_value = float(input_value)
            print("入力された値:", numeric_value)
            return int(input_value)
        except ValueError:
            print("有効な数値を入力してください。")
            return 1

    def to_csv(self):
        ## 参考　https://bunkyudo.co.jp/python-tabula-01/
        ## 複数PDFをCSVに変換(確定申告用)
        ## ヘッダー
        ## お取引日 お取引内容    ご利用額 利用手数料(円)現地通貨額 通貨 レート  Unnamed: 0  Unnamed: 1  確定

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

            dfs = tabula.read_pdf(filename, stream=True, pages='all', java_options="-Dfile.encoding=UTF8")
            count = len(dfs)
            self.add_log(f'{filename} {count} 件')

            if len(dfs) == 0:
                continue

            df = pd.concat(dfs, axis=0)

            base_df = pd.concat([base_df, df[start_row_index:]], axis=0)

        if base_df is not None:
            if not base_df.empty:
                filename = asksaveasfilename(filetype=[('CSV files', '*.csv')])
                if filename:
                    # df.to_csv(filename, header=False, index=False)
                    base_df.to_csv(filename, index=False)
                    self.add_log(f"{filename} にファイルを書き出しました")


# アプリを起動する
if __name__ == "__main__":
    app = PdfToCsvApp()
    app.mainloop()
