#プログラム1｜ライブラリの設定
from datetime import datetime, timedelta
import datetime
import xlwings as xw
import pandas as pd
import os

#プログラム2｜対象エクセルのファイルパスを指定
samplepath = 'list/info.xlsx'
templatepath = 'list/template.xlsx'
#savepath has to be absolute path to use in ExportAsFixedFormat func
# savepath = 'C:/Yamazaki_Rei/misc/BillGenerator/bill/'

#プログラム3｜エクセルを読み込み、日付を変換
df = pd.read_excel(samplepath)
dt_now = datetime.datetime.now()
title = str(dt_now.month + 1) + '月分請求書'
year = str(dt_now.year) + '年_'
fileTitle = year + title
# df['納品日'] = pd.to_datetime(df['納品日']).dt.strftime("%Y-%m-%d")

# Create directories for saving files
current_path = os.getcwd()
new_folder_path = current_path + '/bill/' + fileTitle
new_excel_folder_path = new_folder_path + '/excel/'
os.mkdir(new_folder_path)
os.mkdir(new_excel_folder_path)

#プログラム4｜取引先のリストを作成
torihiki_list = sorted(list(df['生徒氏名'].unique()))

#プログラム5｜エクセルを新しいインスタンスで作成(エクセルのアプリケーションを開く)
App = xw.App()

#プログラム6｜取引先ごとに処理
for torihiki in torihiki_list:

    # プログラム7｜テンプレートエクセルを開く
    wb = App.books.open(templatepath)
    ws = wb.sheets('原本')

    # プログラム8｜対象期間を設定
    # startdate = datetime(int(ws['J4'].value), int(ws['K4'].value), int(ws['L4'].value))
    # enddate = datetime(int(ws['J5'].value), int(ws['K5'].value), int(ws['L5'].value))

    # プログラム9｜情報を設定
    goukei = 0
    gyo1 = 18
    gyo2 = 23
    # プログラム10｜取引先ごとにフィルターしてリストに変換
    filtered = df[df['生徒氏名'] == f'{torihiki}']
    values = filtered.values.tolist()

    # プログラム11｜取引先ごとのデータの内、対象期間に含まれるものだけを処理

    for rows in values:
        print(rows[0])
        ws.range('A1').value = title
        ws.range('J5').value = dt_now.strftime('%Y年%m月%d日')
        reiwa = str(dt_now.year - 2018)
        ws.range('J6').value = reiwa
        ws.range('K6').value = dt_now.month
        ws.range('A2').value = rows[5]
        ws.range('A3').value = rows[6]
        ws.range('A6').value = rows[0]
        ws.range('L6').value = rows[1]
        if(rows[13] != 0):
            titlePos = 'B' + str(gyo1)
            yenPos = 'K' + str(gyo1)
            gyo1 += 1
            ws.range(titlePos).value = '授業料'
            ws.range(yenPos).value = rows[13]
        if(rows[14] != 0):
            titlePos = 'B' + str(gyo1)
            yenPos = 'K' + str(gyo1)
            gyo1 += 1
            ws.range(titlePos).value = '割引'
            ws.range(yenPos).value = rows[14]
        if(rows[15] != 0):
            titlePos = 'B' + str(gyo1)
            yenPos = 'K' + str(gyo1)
            gyo1 += 1
            ws.range(titlePos).value = '施設費'
            ws.range(yenPos).value = rows[15]
        if(rows[16] != 0):
            titlePos = 'B' + str(gyo2)
            yenPos = 'K' + str(gyo2)
            gyo2 += 1
            ws.range(titlePos).value = 'おやつ代'
            ws.range(yenPos).value = rows[16]
        if(rows[17] != 0):
            titlePos = 'B' + str(gyo2)
            yenPos = 'K' + str(gyo2)
            gyo2 += 1
            ws.range(titlePos).value = '送迎'
            ws.range(yenPos).value = rows[17]
        if(rows[18] != 0):
            titlePos = 'B' + str(gyo1)
            yenPos = 'K' + str(gyo1)
            gyo1 += 1
            ws.range(titlePos).value = '延長'
            ws.range(yenPos).value = rows[18]
        if(rows[19] != 0):
            titlePos = 'B' + str(gyo2)
            yenPos = 'K' + str(gyo2)
            gyo2 += 1
            ws.range(titlePos).value = '教材'
            ws.range(yenPos).value = rows[19]
        if(rows[20] != 0):
            titlePos = 'B' + str(gyo2)
            yenPos = 'K' + str(gyo2)
            gyo2 += 1
            ws.range(titlePos).value = 'その他'
            ws.range(yenPos).value = rows[20]
        ws.range('A32').value = rows[11]

    filename = torihiki + '.xlsx'
    # savepath = ws.range('P1').value + '/'
    # saveXlsxPath = ws.range('P2').value + '/'

    pdf_path = os.path.join(new_folder_path, f'{torihiki}_reportXX.pdf')
    wb.api.ExportAsFixedFormat(0, pdf_path)
    os.startfile(pdf_path)

    wb.save(new_excel_folder_path + filename)
    wb.close()

App.quit()
