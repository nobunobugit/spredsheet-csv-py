import gspread
import json
import pandas as pd
import re
import yaml

# ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials

# dataframeをスプレッドシートへインサートする。
def write_spreadsheet(worksheet, df):


    col_lastnum = len(df.columns)  # DataFrameの列数
    row_lastnum = len(df.index)  # DataFrameの行数

    start_cell = 'A2'  # 列はA〜Z列まで
    start_cell_col = re.sub(r'[\d]', '', start_cell)
    start_cell_row = int(re.sub(r'[\D]', '', start_cell))

    # アルファベットから数字を返すラムダ式(A列～Z列まで)
    # 例：A→1、Z→26
    alpha2num = lambda c: ord(c) - ord('A') + 1

    # 展開を開始するセルからA1セルの差分
    row_diff = start_cell_row - 1
    col_diff = alpha2num(start_cell_col) - alpha2num('A')

    # DataFrameのヘッダーと中身をスプレッドシートのC４セルから展開する
    cell_list = worksheet.range(start_cell + ':' + toAlpha(col_lastnum + col_diff) + str(row_lastnum + 1 + row_diff))
    for cell in cell_list:

        if cell.row == 1 + row_diff:
            val = df.columns[cell.col - (1 + col_diff)]

        else:
            val = df.iloc[cell.row - (2 + row_diff)][cell.col - (1 + col_diff)]

        cell.value = val

    worksheet.update_cells(cell_list)


# 数字からアルファベットを返す関数
# 例：26→Z、27→AA、10000→NTP
def toAlpha(num):
    if num <= 26:
        return chr(64 + num)
    elif num % 26 == 0:
        return toAlpha(num // 26 - 1) + chr(90)
    else:
        return toAlpha(num // 26) + chr(64 + num % 26)


def workbook_info(JOSON_PATH, SPREADSHEET_KEY, WORKSHEET_NAME):
    # 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # 認証情報設定
    # ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        JOSON_PATH, scope)

    # OAuth2の資格情報を使用してGoogle APIにログインします。
    gc = gspread.authorize(credentials)

    # 共有設定したスプレッドシートを開く
    return gc.open_by_key(SPREADSHEET_KEY).worksheet(WORKSHEET_NAME)


# 実行部分
if __name__ == "__main__":
    with open('../setting/setting.yml') as file:
        yml = yaml.load(file, Loader=yaml.SafeLoader)

    worksheet = workbook_info(yml["JOSON_PATH"], yml["SPREADSHEET_KEY"], yml["WORKSHEET_NAME"])
    df = pd.read_csv(yml["CSV_FILENAME"])
    write_spreadsheet(worksheet, df.fillna(""))
