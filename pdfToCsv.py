import pandas as pd
import tabula
import os

## 参考　https://bunkyudo.co.jp/python-tabula-01/
## 複数PDFをCSVに変換(確定申告用)
## ヘッダー
## お取引日 お取引内容    ご利用額 利用手数料(円)現地通貨額 通貨 レート  Unnamed: 0  Unnamed: 1  確定

baseDf = None
startRowIndex = 6
baseDir = "../../確定申告/debit/"
index1 = 0

for i in range(2):
    directory = baseDir + str(i + 1) + "/"
    index2 = 0
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):

            dfs = tabula.read_pdf(os.path.join(directory, filename), stream=True, pages='all',
                                  java_options="-Dfile.encoding=UTF8")
            df = pd.concat(dfs, axis=0)

            if index1 + index2 == 0:
                baseDf = df
            else:
                baseDf = pd.concat([baseDf, df[startRowIndex:]], axis=0)
            index2 = index2 + 1
        else:
            continue
    index1 = index1 + 1
if not baseDf.empty:
    baseDf.to_csv("result/result.csv", index=False, encoding='utf-8-sig')
