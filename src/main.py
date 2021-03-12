import datetime
import xlwings as xw

infoPath = '../list/info.xlsx'
templatePath = '../list/template.xlsx'

# Create Title [X月分請求書]
d_now = datetime.datetime.now()
d_month = d_now.month + 1
title = str(d_month) + '月分請求書'

# Create new Excel file
App = xw.App()
wb = xw.Book()

xw.Range('A1').value = 'TEST'
wb.save('test02.xlsx')
