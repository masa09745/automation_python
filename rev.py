import openpyxl as op
import os
import sys
import glob

def lead_excel():
  dir = os.getcwd()
  select_file = ','.join(glob.glob(dir + '/Data/*Rev*.xlsx' ))
  wb = op.load_workbook(select_file)
  worksheets = wb.sheetnames
  ws = wb['RN']
  
  for i in range(ws.max_row, 3, -1):
    if ws.cell(row = i, column = 3).value == 'XLD':
      ws.delete_rows(i)
      wb.save('test.xlsx')
  print (worksheets)
  
  

lead_excel()