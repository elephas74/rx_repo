#!/usr/bin/env python
# coding: utf-8


import pandas as pd
import xlsxwriter
from pandas.io.excel import ExcelWriter

df = pd.read_excel('rxbase.xlsx')

df['dob'] = pd.to_datetime(df['dob'])

df['dob'] = df['dob'].dt.date

df.loc[df.xifo=='x',"rr"] -= 1

df.loc[df.xifo=='x', 'reo']

df['reo'] = pd.to_datetime(df['reo']) 

df['reo'] = df ['reo'].dt.date

df['reo'] = df['reo']+pd.to_timedelta(df['rxday'],unit='d') 

df.loc[df.xifo=='x','xifo'] = " "


writer = pd.ExcelWriter("rxbaseV41.xlsx",
                        engine='xlsxwriter',
                        date_format='mm/dd/yy')

df.to_excel(writer, sheet_name='rxsheet', index=False, na_rep=' ')


workbook = writer.book
worksheet = writer.sheets['rxsheet']

format2 = workbook.add_format({'text_wrap': True})
format3 = workbook.add_format({'center_across': True})

worksheet.set_column('A:A', 25)
worksheet.set_column('B:B', 12, format3)
worksheet.set_column('C:C', 12, format3)
worksheet.set_column('D:D', 12, format3)
worksheet.set_column('E:E', 5, format3)
worksheet.set_column('F:F', 10, format3)
worksheet.set_column('G:G', 5, format3)
worksheet.set_column('H:H', 27, format2)
worksheet.set_column('I:I', 10, format3)
worksheet.set_column('J:K', 20)
worksheet.set_column('L:L', 30, format2)

writer.save()
