import tkinter as tk
import tkinter.messagebox as mb


root = tk.Tk()
root.geometry('300x200')
root.title('作業アサインシート一括印刷')

label1 = tk.Label(text='機番')
label2 = tk.Label(text='エチロン')

label1.place(x=30, y=20)
label2.place(x=30, y=80)


txt1 = tk.Entry(width=20)
txt2 = tk.Entry(width=20)


txt1.place(x=30, y=45)
txt2.place(x=30, y=105)


root.mainloop()