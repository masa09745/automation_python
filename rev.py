import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as mb

from tkinter import messagebox

import openpyxl as op
import os
import sys
import glob

def lead_excel():
  ship_input = ship_text.get()
  ech_input = ech_text.get()
  
  dir = os.getcwd()
  new_dir_filepath = dir + f'/JA{ship_input}_{ech_input}'
  if os.path.exists(new_dir_filepath):
    print('フォルダは存在します')
  else:
    os.mkdir(new_dir_filepath)

  select_file = ','.join(glob.glob(dir + '/Data/*Rev*.xlsx' ))
  wb = op.load_workbook(select_file)
  worksheets = wb.sheetnames
  ws = wb['RN']
  val = ws['A1'].value
  
  for i in range(ws.max_row, 3, -1):
    if ws.cell(row = i, column = 3).value == 'XLD':
      ws.delete_rows(i)
  filename = f'{new_dir_filepath}/{val}.xlsx'
  wb.save(filename)
  
  rev_list = op.load_workbook(f'{new_dir_filepath}/{val}.xlsx')
  word_file = ','.join(glob.glob(dir + 'data/作業アサインシート.docx'))

def click():
  lead_excel()
  messagebox.showinfo("完了", "完了しました")

def close():
  main_win.destroy()
  

main_win = tk.Tk()
main_win.geometry('300x200')
main_win.title('作業アサインシート一括印刷')

main_frame = ttk.Frame(main_win)
main_frame.grid(column=0, row=0, sticky=tk.NSEW, padx=20, pady=20)

ship_label = tk.Label(main_frame, text='機番 (JAは不要)')
ship_text = tk.StringVar()
shiptext_entry = tk.Entry(main_frame, textvariable=ship_text)

ech_label = tk.Label(main_frame, text='エチロン')
ech_text = tk.StringVar()
echtext_entry = tk.Entry(main_frame, textvariable=ech_text)

create_btn = tk.Button(main_frame, text="作成", command=click)
close_btn = tk.Button(main_frame, text="閉じる", command=close)

ship_label.grid(column=0, row=0, sticky=tk.W)
shiptext_entry.grid(column=0, row=1, pady=5, sticky=tk.EW)

ech_label.grid(column=0, row=2, sticky=tk.W)
echtext_entry.grid(column=0, row=3, pady=5, sticky=tk.EW)

create_btn.grid(column=0, row=4, sticky=tk.W, pady=10)
close_btn.grid(column=0, row=4, sticky=tk.E, pady=10)

main_win.columnconfigure(0, weight=1)
main_win.rowconfigure(0, weight=1)
main_frame.columnconfigure(0, weight=1)

main_win.mainloop()