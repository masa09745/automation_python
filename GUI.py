import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as mb
from tkinter import messagebox


import openpyxl as op
import docx as dx
import glob


def create():
  shiptext_value=ship_text.get()
  echtext_value=ech_text.get()
  
  ship_num = shiptext_value
  ech_num = echtext_value

  change_word = [
    ['機番', ship_num],
    ['エチロン', ech_num]
  ]
  
  word_file = glob.glob('必要データ/*.docx')
  select_file = ','.join(word_file)
  doc = dx.Document(select_file)

  tbl = doc.tables[0]
  
  for row in tbl.rows:
    values = []
    for cell in row.cells:
      values.append(cell.text)
      print(cell.text)
  rename_file = f'{ship_num}_{ech_num}.docx'
  doc.save(rename_file)
  messagebox.showinfo("完了", "完了しました")




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

create_btn = tk.Button(main_frame, text="作成", command=create)

ship_label.grid(column=0, row=0, sticky=tk.W)
shiptext_entry.grid(column=0, row=1, pady=5, sticky=tk.EW)

ech_label.grid(column=0, row=3, sticky=tk.W)
echtext_entry.grid(column=0, row=4, pady=5, sticky=tk.EW)

create_btn.grid(column=0, row=5, pady=15)

main_win.columnconfigure(0, weight=1)
main_win.rowconfigure(0, weight=1)
main_frame.columnconfigure(0, weight=1)

main_win.mainloop()