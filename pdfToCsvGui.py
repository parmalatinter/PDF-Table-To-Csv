# PDF TO Excelアプリ
# --------------------------
import glob
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import tabula
import os


class PdfToCsvApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("PDF TO Excel")
        self.geometry("800x400")

        # ラベルとエントリーを配置するフレーム
        input_frame = tk.Frame(self)
        input_frame.pack(pady=10)
        tk.Label(input_frame, text="フォルダ:").grid(row=0, column=0)
        self.task_entry = tk.Entry(input_frame, width=30)
        self.task_entry.grid(row=0, column=1)

        # 追加ボタンを配置するフレーム
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)

        # タスクの一覧を表示するリストボックス
        self.task_list = tk.Listbox(self, selectmode="multiple", bd=10)
        self.task_list.pack(fill="both", expand=True)

        # ボタンを作成してフォルダ選択関数を呼び出す
        tk.Button(btn_frame, text="フォルダを選択", command=self.select_folder).pack(pady=10)
        tk.Button(btn_frame, text="To CSV", command=self.to_csv).pack(side="left")

        # タスクの一覧
        self.tasks = []

    # タスクの追加
    def add_log(self, msg):
        self.tasks.append(msg)
        self.task_list.insert("end", msg)
        self.task_entry.delete(0, "end")

    # タスクの削除
    def delete_log(self):
        self.task_list.delete(0, tk.END)

    # Todoアプリのクラス
    def select_folder(self):
        folder_path = filedialog.askdirectory()  # フォルダ選択ダイアログを表示
        if folder_path:
            print("選択したフォルダ:", folder_path)
            self.task_entry.insert(0, folder_path)

    def to_csv(self):
        ## 参考　https://bunkyudo.co.jp/python-tabula-01/
        ## 複数PDFをCSVに変換(確定申告用)
        ## ヘッダー
        ## お取引日 お取引内容    ご利用額 利用手数料(円)現地通貨額 通貨 レート  Unnamed: 0  Unnamed: 1  確定

        base_df = None
        start_row_index = 1
        base_dir = self.task_entry.get()
        self.delete_log()
        for filename in glob.glob(base_dir + '/*pdf'):


            dfs = tabula.read_pdf(filename, stream=True, pages='all',
                                  java_options="-Dfile.encoding=UTF8")
            count = len(dfs)
            self.add_log(f'{filename} {count} 件')
            if len(dfs) == 0:
                continue

            df = pd.concat(dfs, axis=0)

            base_df = pd.concat([base_df, df[start_row_index:]], axis=0)

        if base_df is not None:
            if not base_df.empty:
                base_df.to_csv(f"{base_dir}/result.csv", index=False, encoding='utf-8-sig')
                self.add_log(f"{base_dir}/result.csv にファイルを書き出しました")


# PDF TO Excelアプリを起動する
if __name__ == "__main__":
    app = PdfToCsvApp()
    app.mainloop()
