import datetime
import xlwings as xw
import pandas as pd

infoPath = '../list/info.xlsx'
templatePath = '../list/template.xlsx'

# Create Title [X月分請求書]
d_now = datetime.datetime.now()
d_month = d_now.month + 1
title = str(d_month) + '月分請求書'

# Create new Excel file
# App = xw.App()
# wb = xw.Book()

# Write data into Excel file
# xw.Range('A1').value = title
# wb.save('test.xlsx')

df = pd.read_excel(infoPath)
print("read_excel succeeded")
